"""Playwright adapter. Selectors are configurable because the live site cannot be inspected here."""
import io
import json
import os
from pathlib import Path
import shutil

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

URL = "https://freecheckprint.com/print"
CONFIG = Path(__file__).with_name("selectors.json")

class BrowserAutomationError(RuntimeError):
    pass

def _find(page, choices, timeout=1500):
    for selector in choices:
        try:
            locator = page.locator(selector).first
            if locator.count() and locator.is_visible(timeout=timeout):
                return locator
        except Exception:
            pass
    return None

def _fill(page, cfg, field, value, required=True):
    locator = _find(page, cfg.get("fields", {}).get(field, []))
    if locator is None:
        if required:
            raise BrowserAutomationError(f"无法定位网站字段 {field}；请检查 selectors.json 并参考页面截图。")
        return
    locator.fill(str(value))

def generate_check(data, accept_terms=False):
    if not accept_terms:
        raise BrowserAutomationError("Terms must be explicitly accepted by the user.")
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    screenshot = None
    with sync_playwright() as p:
        chromium_path = shutil.which("chromium") or shutil.which("chromium-browser")

        if not chromium_path:
            raise RuntimeError(
                "Chromium not found. Please add 'chromium' to packages.txt."
            )
        
        browser = p.chromium.launch(
            headless=True,
            executable_path=chromium_path,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(accept_downloads=True, viewport={"width":1280,"height":1000})
        page = context.new_page()
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=45000)
            for field in ("company","address","bank","routing","account","check_number","date","payee","amount","memo"):
                _fill(page,cfg,field,data.get(field,""),required=field not in ("address","memo","bank"))
            # Check Stock only: never silently fall back to Plain Paper.
            stock = _find(page,cfg["check_stock"] )
            if stock is None:
                raise BrowserAutomationError("找不到 Check Stock 选项，已停止（不会使用 Plain Paper）。")
            stock.click()
            terms = _find(page,cfg["terms"] )
            if terms is None:
                raise BrowserAutomationError("找不到 Terms checkbox，已停止。")
            if terms.get_attribute("type") == "checkbox":
                terms.check()
            else:
                terms.click()
            page.screenshot(path="/tmp/fcp_before_submit.png", full_page=True)
            submit = _find(page,cfg["generate"] )
            if submit is None:
                raise BrowserAutomationError("找不到 Generate/Print 按钮，已停止。")
            # The site might download, open a new tab, or use the browser print dialog.
            pdf = None
            try:
                with page.expect_download(timeout=10000) as dl:
                    submit.click()
                download = dl.value
                path = download.path()
                if path and path.lower().endswith(".pdf"):
                    pdf = Path(path).read_bytes()
                elif path:
                    raw=Path(path).read_bytes()
                    if raw.startswith(b"%PDF"):
                        pdf=raw
            except PlaywrightTimeout:
                pass
            if pdf is None:
                # Do not use page.pdf(): it may print the web form rather than the check.
                # Search for an actual PDF in any new tab or embedded object.
                for tab in context.pages:
                    try:
                        if tab.url.lower().endswith(".pdf"):
                            resp=context.request.get(tab.url)
                            if resp.ok and resp.body().startswith(b"%PDF"):
                                pdf=resp.body();break
                    except Exception:
                        pass
            screenshot=page.screenshot(full_page=True)
            if pdf is None:
                return {"pdf":None,"check_number":data["check_number"],"screenshot":screenshot,
                        "message":"表单可能已提交，但没有检测到网站的 PDF 下载。请根据截图调整选择器/下载流程；不会把网页直接当支票 PDF。"}
            return {"pdf":pdf,"check_number":data["check_number"],"screenshot":screenshot}
        except BrowserAutomationError:
            raise
        except Exception as e:
            raise BrowserAutomationError(f"FreeCheckPrint 自动化失败：{type(e).__name__}: {e}") from e
        finally:
            browser.close()
