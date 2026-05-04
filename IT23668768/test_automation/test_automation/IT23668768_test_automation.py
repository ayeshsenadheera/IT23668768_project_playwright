from playwright.sync_api import sync_playwright
import openpyxl
import argparse
import time

def dismiss_popup(page):
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
    except:
        pass
    try:
        overlay = page.locator("div.fixed.inset-0")
        if overlay.count() > 0 and overlay.first.is_visible():
            page.mouse.click(5, 5)
            page.wait_for_timeout(500)
    except:
        pass

def run_tests(excel_path, url, wait_ms, slow_mo_ms):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Find columns from header row
    input_col = expected_col = actual_col = status_col = None
    for cell in ws[1]:
        if cell.value:
            v = str(cell.value).strip().lower()
            if v == "input":
                input_col = cell.column
            elif "expected" in v:
                expected_col = cell.column
            elif "actual" in v:
                actual_col = cell.column
            elif v == "status":
                status_col = cell.column

    print(f"Input col: {input_col}, Expected col: {expected_col}, Actual col: {actual_col}, Status col: {status_col}")

    if not input_col:
        print("ERROR: Cannot find Input column!")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=slow_mo_ms)
        page = browser.new_page()
        page.set_default_timeout(30000)

        print(f"Opening {url} ...")
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)

        # Dismiss any startup popup
        dismiss_popup(page)
        page.wait_for_timeout(1000)

        # Find input box
        print("Finding input box...")
        input_box = page.locator('textarea[placeholder*="Singlish"]').first
        input_box.wait_for(state="visible", timeout=10000)
        print("Input box found!")

        # Find translate button
        translate_btn = page.locator('button:has-text("Translate")').first
        print("Translate button found!")

        # Find output div
        output_div = page.locator('div.whitespace-pre-wrap').first
        print("Output div found!")
        print("Starting tests...\n")

        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=input_col).value
            if not cell_val:
                continue

            singlish = str(cell_val).strip()
            expected = str(ws.cell(row=row, column=expected_col).value).strip() if expected_col else ""

            print(f"Testing [Row {row}]: {singlish}")

            try:
                # Dismiss any popups first
                dismiss_popup(page)

                # Clear input using JavaScript
                page.evaluate(
                    """(selector) => {
                        const el = document.querySelector(selector);
                        if (el) { el.value = ''; el.dispatchEvent(new Event('input', {bubbles:true})); }
                    }""",
                    'textarea[placeholder*="Singlish"]'
                )
                page.wait_for_timeout(300)

                # Type singlish text using JavaScript fill
                input_box.click(force=True)
                page.wait_for_timeout(300)
                input_box.fill(singlish)
                page.wait_for_timeout(500)

                # Click translate
                translate_btn.click(force=True)
                page.wait_for_timeout(wait_ms)

                # Read output
                actual = ""
                for _ in range(5):
                    try:
                        actual = output_div.inner_text().strip()
                        if actual:
                            break
                    except:
                        pass
                    page.wait_for_timeout(1000)

                # Determine status
                if actual:
                    status = "PASS" if actual == expected else "FAIL"
                else:
                    status = "UI Error"

                # Save to Excel
                if actual_col:
                    ws.cell(row=row, column=actual_col).value = actual
                if status_col:
                    ws.cell(row=row, column=status_col).value = status
                wb.save(excel_path)

                print(f"  Actual: {actual[:60] if actual else '(empty)'}")
                print(f"  -> {status}\n")

            except Exception as e:
                print(f"  ERROR: {e}\n")
                if status_col:
                    ws.cell(row=row, column=status_col).value = "UI Error"
                wb.save(excel_path)

        print("All tests complete!")
        wb.save(excel_path)
        print(f"Results saved to: {excel_path}")
        print("Keeping browser open. Press CTRL+C to stop.")

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass

        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True)
    parser.add_argument("--url", default="https://www.pixelssuite.com/transliteration")
    parser.add_argument("--wait-ms", type=int, default=8000)
    parser.add_argument("--type-delay-ms", type=int, default=100)
    parser.add_argument("--slow-mo-ms", type=int, default=300)
    parser.add_argument("--save-every", type=int, default=1)
    parser.add_argument("--keep-open", action="store_true", default=False)
    args = parser.parse_args()

    run_tests(
        excel_path=args.excel,
        url=args.url,
        wait_ms=args.wait_ms,
        slow_mo_ms=args.slow_mo_ms,
    )
