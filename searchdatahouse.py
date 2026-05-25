# -*- coding: utf-8 -*-
"""Portal Inmobiliario (Portalinmobiliario.com) scraper

- Selenium-only (passes JS challenge)
- Parses Nordic JSON from __NORDIC_RENDERING_CTX__ / _n.ctx.r
- Paginates using pagination.next_page.url (pattern _Desde_49_ ...)
- Exports formatted Excel (no #######)
- Avoids PermissionError by saving to a unique filename when locked
- Checkpointing every N rows

Usage:
  pip install -r requirements.txt
  python searchdatahouse.py

Edit MAX_LISTINGS to control how many listings to collect.
"""

import json
import os
import random
import re
import time
from datetime import datetime, date
from pathlib import Path

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIG
# ============================================================
LISTING_URL = (
    "https://www.portalinmobiliario.com/venta/casa/propiedades-usadas/"
    "las-condes-metropolitana/_OrderId_PRICE_PriceRange_0CLP-400000000CLP_NoIndex_True"
)

CITY = "las_condes"
MAX_LISTINGS = 300          # prueba intermedia; sube a 100000 para traer "todas"
# Local: ventana visible. En GitHub Actions se pone HEADLESS=true (ver env del workflow)
HEADLESS = os.environ.get("HEADLESS", "false").lower() == "true"
CHECKPOINT_EVERY_N = 5

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/135.0 Safari/537.36"
)

DEBUG_HTML_FILE = "debug_selenium.html"
CHECKPOINT_FILE = f"checkpoint_{CITY}_{date.today().isoformat()}.json"
OUTPUT_XLSX = f"properties_{CITY}_{date.today().isoformat()}.xlsx"


# ============================================================
# UTILS
# ============================================================
def iso_now():
    return datetime.now().isoformat(timespec="seconds")


def polite_sleep(a=1.0, b=3.0):
    time.sleep(random.uniform(a, b))


def clean_number(text):
    if not text:
        return None
    digits = re.sub(r"[^\d]", "", str(text))
    return int(digits) if digits else None


def normalize_url(u):
    if not u:
        return ""
    if u.startswith("http"):
        return u
    return "https://www." + u.lstrip("/")


def detect_block_or_captcha(html: str) -> bool:
    if not html:
        return False
    s = html.lower()
    keywords = [
        "captcha", "verifica que eres", "verify you are", "robot",
        "soy humano", "challenge", "security check"
    ]
    return any(k in s for k in keywords)


def load_checkpoint():
    p = Path(CHECKPOINT_FILE)
    if not p.exists():
        return {
            "page_number": 1,
            "current_url": LISTING_URL,
            "completed_rows": 0,
            "scraped_urls": [],
            "rows": [],
            "errors": []
        }
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {
            "page_number": 1,
            "current_url": LISTING_URL,
            "completed_rows": 0,
            "scraped_urls": [],
            "rows": [],
            "errors": []
        }


def save_checkpoint(state):
    Path(CHECKPOINT_FILE).write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


# ============================================================
# FILE LOCK SAFE OUTPUT (fix PermissionError)
# ============================================================
def unique_name(path: str) -> str:
    p = Path(path)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return str(p.with_name(f"{p.stem}_{stamp}{p.suffix}"))


def safe_excel_path(path: str) -> str:
    p = Path(path)
    if not p.exists():
        return str(p)

    try:
        tmp = p.with_suffix(p.suffix + ".locktest")
        p.rename(tmp)
        tmp.rename(p)
        return str(p)
    except PermissionError:
        return unique_name(path)


# ============================================================
# OVERLAY DISMISS
# ============================================================
def dismiss_overlays(driver, tries=4):
    xpaths = [
        "//button[contains(., 'Entendido')]",
        "//button[contains(., 'Aceptar')]",
        "//button[contains(., 'Acepto')]",
        "//button[contains(., 'Aceptar todo')]",
        "//button[contains(., 'Aceptar todas')]",
        "//button[contains(., 'Accept')]",
        "//button[contains(., 'Accept all')]",
        "//button[contains(., 'OK')]",
    ]
    for _ in range(tries):
        clicked = False
        for xp in xpaths:
            try:
                btns = driver.find_elements(By.XPATH, xp)
                for b in btns:
                    try:
                        if b.is_displayed() and b.is_enabled():
                            b.click()
                            clicked = True
                            polite_sleep(0.15, 0.4)
                    except Exception:
                        pass
            except Exception:
                pass
        if not clicked:
            break


# ============================================================
# NORDIC JSON
# ============================================================
def wait_for_nordic_ctx(driver, timeout=90):
    wait = WebDriverWait(driver, timeout)
    wait.until(lambda d: len(d.find_elements(By.ID, "__NORDIC_RENDERING_CTX__")) > 0)

    def has_marker(d):
        try:
            el = d.find_element(By.ID, "__NORDIC_RENDERING_CTX__")
            txt = el.get_attribute("textContent") or ""
            return "_n.ctx.r=" in txt
        except Exception:
            return False

    wait.until(has_marker)


def extract_nordic_json_from_dom(driver):
    script = driver.find_element(By.ID, "__NORDIC_RENDERING_CTX__")
    txt = script.get_attribute("textContent") or ""

    marker = "_n.ctx.r="
    idx = txt.find(marker)
    if idx == -1:
        return None

    payload = txt[idx + len(marker):].lstrip()
    decoder = json.JSONDecoder()
    obj, end = decoder.raw_decode(payload)
    return obj


def parse_polycards(nordic_data):
    results = (
        nordic_data.get("appProps", {})
                  .get("pageProps", {})
                  .get("initialState", {})
                  .get("results", [])
    )
    return [x for x in results if x.get("id") == "POLYCARD"]


def components_to_dict(components):
    out = {}
    for c in components or []:
        t = c.get("type")
        if t:
            out[t] = c.get(t, {})
    return out


def get_next_page_url(nordic_data):
    return (
        nordic_data.get("appProps", {})
                  .get("pageProps", {})
                  .get("initialState", {})
                  .get("pagination", {})
                  .get("next_page", {})
                  .get("url")
    )


# ============================================================
# EXCEL FORMATTING
# ============================================================
def format_excel(file_path: str, sheet_name="Listings"):
    wb = openpyxl.load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.save(file_path)
        return

    ws = wb[sheet_name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}

    wrap_cols = {"Title", "Address", "Seller", "Description", "Source_URL"}
    for name in wrap_cols:
        if name in headers:
            col = headers[name]
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")

    if "Price_CLP" in headers:
        col = headers["Price_CLP"]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '"$"#,##0'

    if "Price_UF" in headers:
        col = headers["Price_UF"]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '"UF"#,##0'

    if "Price_USD" in headers:
        col = headers["Price_USD"]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '"US$"#,##0.00'

    alt_fill = PatternFill("solid", fgColor="F7F7F7")
    for r in range(2, ws.max_row + 1):
        if r % 2 == 0:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = alt_fill

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            v = ws.cell(row_idx, col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    wb.save(file_path)


# ============================================================
# MAIN
# ============================================================
def main():
    state = load_checkpoint()

    page_number = state.get("page_number", 1)
    current_url = state.get("current_url", LISTING_URL) or LISTING_URL

    scraped_urls = set(state.get("scraped_urls", []))
    rows = state.get("rows", [])
    errors = state.get("errors", [])

    duplicates_skipped = 0

    chrome_options = Options()
    if HEADLESS:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument(f"--user-agent={USER_AGENT}")

    driver = webdriver.Chrome(options=chrome_options)

    try:
        while len(rows) < MAX_LISTINGS and current_url:
            driver.get(current_url)
            polite_sleep(1.0, 2.0)
            dismiss_overlays(driver)

            wait_for_nordic_ctx(driver, timeout=90)
            polite_sleep(0.7, 1.5)
            dismiss_overlays(driver)

            html = driver.page_source
            Path(DEBUG_HTML_FILE).write_text(html, encoding="utf-8")

            if detect_block_or_captcha(html):
                driver.save_screenshot("captcha_or_block.png")
                raise SystemExit("Bloqueo/CAPTCHA detectado. Screenshot: captcha_or_block.png")

            nordic_data = extract_nordic_json_from_dom(driver)
            if not nordic_data:
                raise SystemExit("Nordic JSON no encontrado. Revisa debug_selenium.html")

            polycards = parse_polycards(nordic_data)

            for item in polycards:
                if len(rows) >= MAX_LISTINGS:
                    break

                poly = item.get("polycard", {})
                meta = poly.get("metadata", {})
                comps = components_to_dict(poly.get("components", []))

                url = normalize_url(meta.get("url", ""))

                if url and url in scraped_urls:
                    duplicates_skipped += 1
                    continue

                headline = (comps.get("headline", {}).get("text") or "Not listed").strip()
                title = (comps.get("title", {}).get("text") or "Not listed").strip()
                address = (comps.get("location", {}).get("text") or "Not listed").strip()
                seller = (comps.get("seller", {}).get("text") or "Not listed").strip()

                title = re.sub(r"\s+", " ", title).strip()
                seller = re.sub(r"\s+", " ", seller).strip()

                if "departamento" in title.lower():
                    continue
                if "casa" not in headline.lower():
                    continue

                price_block = comps.get("price", {}).get("current_price", {}) or {}
                price_value = price_block.get("value")
                price_currency = price_block.get("currency")

                price_clp = price_value if price_currency == "CLP" else None
                price_uf = price_value if price_currency == "CLF" else None
                price_usd = price_value if price_currency == "USD" else None

                attrs = comps.get("attributes_list", {}).get("texts", []) or []
                bedrooms = bathrooms = useful_area = None
                for t in attrs:
                    low = t.lower()
                    if "dorm" in low:
                        bedrooms = clean_number(low)
                    elif "bañ" in low:
                        bathrooms = clean_number(low)
                    elif "m²" in low or "m2" in low:
                        useful_area = clean_number(low)

                rows.append({
                    "Headline": headline,
                    "Title": title,
                    "Address": address,
                    "Comuna": "Las Condes",
                    "Price_CLP": price_clp,
                    "Price_UF": price_uf,
                    "Price_USD": price_usd,
                    "Bedrooms": bedrooms if bedrooms is not None else "Not listed",
                    "Bathrooms": bathrooms if bathrooms is not None else "Not listed",
                    "Useful_Area_m2": useful_area if useful_area is not None else "Not listed",
                    "Seller": seller,
                    "Description": "Not listed",
                    "Source_URL": url or "Not listed",
                    "Scraped_At": iso_now(),
                })

                if url:
                    scraped_urls.add(url)

                if len(rows) % CHECKPOINT_EVERY_N == 0:
                    save_checkpoint({
                        "page_number": page_number,
                        "current_url": current_url,
                        "completed_rows": len(rows),
                        "scraped_urls": sorted(scraped_urls),
                        "rows": rows,
                        "errors": errors
                    })

                polite_sleep(1.0, 3.0)

            next_url = get_next_page_url(nordic_data)
            if next_url:
                current_url = next_url
                page_number += 1
                polite_sleep(2.0, 5.0)
            else:
                current_url = None

            save_checkpoint({
                "page_number": page_number,
                "current_url": current_url if current_url else "",
                "completed_rows": len(rows),
                "scraped_urls": sorted(scraped_urls),
                "rows": rows,
                "errors": errors
            })

    finally:
        driver.quit()

    df = pd.DataFrame(rows)
    err_df = pd.DataFrame(errors)

    columns = [
        "Headline", "Title", "Address", "Comuna",
        "Price_CLP", "Price_UF", "Price_USD",
        "Bedrooms", "Bathrooms", "Useful_Area_m2",
        "Seller", "Description", "Source_URL", "Scraped_At"
    ]

    for c in columns:
        if c not in df.columns:
            df[c] = "Not listed"
    df = df[columns]

    out_path = safe_excel_path(OUTPUT_XLSX)

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Listings")
            if not err_df.empty:
                err_df.to_excel(writer, index=False, sheet_name="Errors")
            else:
                pd.DataFrame([{
                    "Timestamp": "",
                    "Listing URL": "",
                    "Error Type": "",
                    "Error Message": "",
                    "Retry Count": ""
                }]).to_excel(writer, index=False, sheet_name="Errors")

            pd.DataFrame([{
                "Listing URL": LISTING_URL,
                "City": CITY,
                "Max Listings": MAX_LISTINGS,
                "Headless": HEADLESS,
                "Scraped At": iso_now(),
                "Debug HTML": DEBUG_HTML_FILE,
                "Checkpoint": CHECKPOINT_FILE,
                "Last Page Number": page_number,
                "Collected Rows": len(df),
                "Duplicates Skipped": duplicates_skipped
            }]).to_excel(writer, index=False, sheet_name="Session Info")

        format_excel(out_path, sheet_name="Listings")

    except PermissionError:
        alt = unique_name(OUTPUT_XLSX)
        with pd.ExcelWriter(alt, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Listings")
            if not err_df.empty:
                err_df.to_excel(writer, index=False, sheet_name="Errors")
            pd.DataFrame([{
                "Listing URL": LISTING_URL,
                "City": CITY,
                "Max Listings": MAX_LISTINGS,
                "Headless": HEADLESS,
                "Scraped At": iso_now(),
                "Debug HTML": DEBUG_HTML_FILE,
                "Checkpoint": CHECKPOINT_FILE,
                "Last Page Number": page_number,
                "Collected Rows": len(df),
                "Duplicates Skipped": duplicates_skipped
            }]).to_excel(writer, index=False, sheet_name="Session Info")

        format_excel(alt, sheet_name="Listings")
        out_path = alt

    print(f"✓ {len(df)} properties collected (unique)")
    print(f"✗ {len(errors)} errors (see Errors tab)")
    print(f"⏭ {duplicates_skipped} duplicates skipped")
    print(f"File ready: {out_path}")
    print(f"Checkpoint saved: {CHECKPOINT_FILE}")
    print(f"Debug HTML saved: {DEBUG_HTML_FILE}")


if __name__ == "__main__":
    main()
