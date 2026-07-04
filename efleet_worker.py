from __future__ import annotations

import os
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

from playwright.sync_api import Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError, sync_playwright

from .control_hub import ControlHub
from .excel_manager import ExcelManager


class EfleetWorker:
    def __init__(self, cfg: Dict[str, Any], control: ControlHub, excel: ExcelManager):
        self.cfg = cfg
        self.control = control
        self.excel = excel
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.current_detail_page = None
        self.current_image_page = None
        self.static_dir = Path("static")
        self.static_dir.mkdir(exist_ok=True)

    def run(self) -> None:
        self.control.update_status(running=True, message="Starting bot")
        try:
            self.excel.open()
            rows = self.excel.iter_data_rows(self.cfg.get("start_excel_row", 2))
            self.control.update_status(total_rows=len(rows), message=f"Loaded {len(rows)} rows from Excel")
            if not rows:
                self.control.update_status(message="No HM rows found", running=False)
                return

            if not self.cfg.get("auto_start", False):
                self.control.update_status(paused=True, message="Ready. Press START from control panel.")
                self._wait_while_paused()

            self._setup_browser()
            self._login()

            for idx, (excel_row, hm_no) in enumerate(rows, start=1):
                if self.control.should_stop():
                    break
                self.control.clear_runtime_flags_for_next_row()
                self.control.update_status(
                    current_index=idx,
                    current_excel_row=excel_row,
                    current_hm=hm_no,
                    current_vehicle=hm_no,
                    message=f"Processing HM {hm_no} ({idx}/{len(rows)})",
                    last_error="",
                )
                self.excel.set_row_status(excel_row, "PROCESSING")
                self.excel.apply_defaults_to_row(excel_row)

                try:
                    while True:
                        self._wait_while_paused()
                        self._process_hm(excel_row, hm_no)

                        if self.control.pop_command("REOPEN_HM"):
                            self.control.update_status(message=f"Reopening same HM {hm_no}")
                            self._close_detail_page()
                            continue
                        break

                    self._wait_for_save_next_or_skip(excel_row, hm_no)

                except Exception as e:
                    err = f"{type(e).__name__}: {e}"
                    self.excel.set_row_status(excel_row, "ERROR", err[:250])
                    self.control.update_status(last_error=err, message=f"Error on HM {hm_no}: {err[:120]}")
                    traceback.print_exc()
                finally:
                    self._close_image_page()
                    self._close_detail_page()
                    self._return_to_dashboard()

            self.control.update_status(message="Automation completed", running=False, paused=True)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            self.control.update_status(last_error=err, message=f"Fatal error: {err}", running=False, paused=True)
            traceback.print_exc()
        finally:
            self._cleanup()

    def _setup_browser(self) -> None:
        self.control.update_status(message="Opening Chromium")
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=bool(self.cfg.get("headless", True)),
            args=["--disable-notifications", "--disable-blink-features=AutomationControlled"],
        )
        self.context = self.browser.new_context(
            viewport={"width": 1366, "height": 768},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        self.page = self.context.new_page()

    def _login(self) -> None:
        url = self.cfg.get("website_url")
        username = self.cfg.get("username")
        password = self.cfg.get("password")
        if not url or not username or not password:
            raise ValueError("Missing website_url/username/password. Set EFLEET_URL, EFLEET_USERNAME, EFLEET_PASSWORD or config.json.")
        selectors = self.cfg["selectors"]
        self.control.update_status(message="Logging in")
        self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
        self.page.locator(selectors["login_username"]).fill(username)
        self.page.locator(selectors["login_password"]).fill(password)
        self.page.locator(selectors["login_button"]).click()
        self.page.wait_for_selector(selectors.get("dashboard_ready", "input[type='search']"), timeout=60000)
        self.control.update_status(message="Login successful")

    def _process_hm(self, excel_row: int, hm_no: str) -> None:
        self._search_hm(hm_no)
        self.current_detail_page = self._open_first_view_result(hm_no)
        self.current_detail_page.bring_to_front()
        self._click_first_available(self.current_detail_page, self.cfg["selectors"].get("step2_selectors", []), "Step 2")
        self._small_wait(0.8)
        self._click_first_available(self.current_detail_page, self.cfg["selectors"].get("image_step_selectors", []), "Images step")
        self._small_wait(1.0)
        self._view_images_loop(excel_row)

    def _search_hm(self, hm_no: str) -> None:
        self.control.update_status(message=f"Searching HM {hm_no}")
        search_selector = self.cfg["selectors"].get("search_box", "input[type='search']")
        search = self.page.locator(search_selector).first
        search.wait_for(state="visible", timeout=30000)
        search.click()
        search.fill("")
        search.fill(hm_no)
        # E-fleet search tables often filter on Enter or Tab; use both gently.
        search.press("Enter")
        self._small_wait(1.0)
        search.press("Tab")
        self._small_wait(0.8)

    def _open_first_view_result(self, hm_no: str):
        selectors: List[str] = self.cfg["selectors"].get("eye_selectors", [])
        last_error = ""
        for selector in selectors:
            try:
                locator = self.page.locator(selector).first
                if locator.count() == 0:
                    continue
                locator.scroll_into_view_if_needed(timeout=5000)
                self.control.update_status(message=f"Opening view for HM {hm_no}")
                try:
                    with self.context.expect_page(timeout=8000) as page_info:
                        locator.click(modifiers=["Control"], timeout=5000)
                    detail = page_info.value
                    detail.wait_for_load_state("domcontentloaded", timeout=30000)
                    return detail
                except PlaywrightTimeoutError:
                    # Fallback: focus then Ctrl+Enter, useful when the span itself is not directly clickable.
                    locator.focus(timeout=3000)
                    with self.context.expect_page(timeout=8000) as page_info:
                        self.page.keyboard.press("Control+Enter")
                    detail = page_info.value
                    detail.wait_for_load_state("domcontentloaded", timeout=30000)
                    return detail
            except Exception as e:
                last_error = str(e)[:200]
                continue
        raise RuntimeError(f"Could not open view/eye result for HM {hm_no}. Last error: {last_error}")

    def _click_first_available(self, page, selectors: List[str], label: str) -> bool:
        for selector in selectors:
            try:
                loc = page.locator(selector).first
                if loc.count() > 0:
                    loc.scroll_into_view_if_needed(timeout=5000)
                    loc.click(timeout=7000)
                    self.control.update_status(message=f"Clicked {label}")
                    return True
            except Exception:
                continue
        self.control.update_status(message=f"{label} not found, continuing")
        return False

    def _view_images_loop(self, excel_row: int) -> None:
        page = self.current_detail_page
        selectors = self.cfg["selectors"].get("image_link_selectors", [])
        links = None
        for selector in selectors:
            try:
                links = page.locator(selector)
                count = links.count()
                if count > 0:
                    break
            except Exception:
                links = None
        if links is None or links.count() == 0:
            self.control.update_status(total_images=0, message="No images found. Waiting for Save & Next.")
            return

        total = min(links.count(), int(self.cfg.get("max_images", 20)))
        self.control.update_status(total_images=total, message=f"Found {total} images")
        for i in range(total):
            if self.control.should_stop() or self.control.has_command("STOP_IMAGES", "SAVE_NEXT", "SKIP"):
                break
            self._wait_while_paused()
            self.control.update_status(current_image_index=i + 1, total_images=total, message=f"Opening image {i + 1}/{total}")
            self._open_and_capture_image(links.nth(i), i + 1, total)

    def _open_and_capture_image(self, link, image_no: int, total: int) -> None:
        self._close_image_page()
        try:
            try:
                with self.context.expect_page(timeout=5000) as img_info:
                    link.scroll_into_view_if_needed(timeout=5000)
                    link.click(modifiers=["Control"], timeout=5000)
                self.current_image_page = img_info.value
            except PlaywrightTimeoutError:
                href = link.get_attribute("href")
                if not href:
                    raise
                self.current_image_page = self.context.new_page()
                self.current_image_page.goto(href, wait_until="domcontentloaded", timeout=30000)

            self.current_image_page.wait_for_load_state("domcontentloaded", timeout=30000)
            self.current_image_page.bring_to_front()
            image_path = self.static_dir / "current_image.png"
            try:
                self.current_image_page.screenshot(path=str(image_path), full_page=True, timeout=10000)
                self.control.update_status(current_image_path="/static/current_image.png", message=f"Image {image_no}/{total} captured")
            except Exception as ss_err:
                self.control.update_status(message=f"Image {image_no}/{total} opened; screenshot failed: {str(ss_err)[:80]}")

            start = time.time()
            auto_close = float(self.cfg.get("image_auto_close_sec", 2.0))
            while time.time() - start < auto_close:
                if self.control.should_stop() or self.control.pop_command("CLOSE_IMAGE") or self.control.has_command("STOP_IMAGES", "SAVE_NEXT", "SKIP"):
                    break
                self._wait_while_paused()
                time.sleep(0.1)
        except Exception as e:
            self.control.update_status(last_error=str(e), message=f"Image {image_no} failed, skipped")
        finally:
            self._close_image_page()

    def _wait_for_save_next_or_skip(self, excel_row: int, hm_no: str) -> None:
        if not self.cfg.get("wait_for_save_next_after_images", True):
            self.excel.set_row_status(excel_row, "DONE")
            return
        self.control.update_status(message=f"HM {hm_no} ready. Apply queries, then press Save & Next / Skip.")
        while not self.control.should_stop():
            self._wait_while_paused()
            if self.control.pop_command("SAVE_NEXT"):
                self.excel.set_row_status(excel_row, "DONE")
                self.control.update_status(message=f"HM {hm_no} saved. Moving next.")
                return
            if self.control.pop_command("SKIP"):
                self.excel.set_row_status(excel_row, "SKIPPED")
                self.control.update_status(message=f"HM {hm_no} skipped. Moving next.")
                return
            if self.control.has_command("REOPEN_HM"):
                return
            time.sleep(0.2)

    def _wait_while_paused(self) -> None:
        while self.control.is_paused() and not self.control.should_stop():
            time.sleep(0.2)

    def _small_wait(self, seconds: float) -> None:
        end = time.time() + seconds
        while time.time() < end:
            if self.control.should_stop():
                break
            self._wait_while_paused()
            time.sleep(0.05)

    def _close_image_page(self) -> None:
        try:
            if self.current_image_page and not self.current_image_page.is_closed():
                self.current_image_page.close()
        except Exception:
            pass
        self.current_image_page = None

    def _close_detail_page(self) -> None:
        try:
            if self.current_detail_page and not self.current_detail_page.is_closed():
                self.current_detail_page.close()
        except Exception:
            pass
        self.current_detail_page = None

    def _return_to_dashboard(self) -> None:
        try:
            if self.page and not self.page.is_closed():
                self.page.bring_to_front()
                search_selector = self.cfg["selectors"].get("search_box", "input[type='search']")
                s = self.page.locator(search_selector).first
                if s.count() > 0:
                    s.click(timeout=3000)
                    s.fill("")
        except Exception:
            pass

    def _cleanup(self) -> None:
        self._close_image_page()
        self._close_detail_page()
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.control.update_status(running=False)
