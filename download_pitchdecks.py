"""
SlideShare Pitch Deck PDF Downloader
====================================
Downloads SlideShare presentations as high-resolution PDFs from an Excel/CSV sheet.
Naming format: {Company_Name}.pdf (as given in the spreadsheet).
"""

import os
import sys
import re
import io
import time
import argparse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# Ensure UTF-8 output on Windows consoles to prevent charmap encoding errors
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

import requests
import pandas as pd
from PIL import Image
from rich.console import Console
from rich.table import Table
from playwright.sync_api import sync_playwright

console = Console(force_terminal=True, legacy_windows=False)

# Quality fallback pairs: (quality, width)
RESOLUTIONS = [
    (75, 2048),
    (95, 1024),
    (85, 638),
    (85, 320),
    (95, 2048),
    (85, 2048),
]


def sanitize_filename(name: str) -> str:
    """Sanitize company name to create a safe Windows file name."""
    if not isinstance(name, str):
        name = str(name)
    sanitized = re.sub(r'[\\/*?:"<>|]', '_', name.strip())
    sanitized = re.sub(r'[\s_]+', '_', sanitized)
    sanitized = sanitized.strip('. ')
    return sanitized if sanitized else "Untitled"


def find_columns(df: pd.DataFrame, custom_name_col: str = None, custom_url_col: str = None):
    """Auto-detect Company Name and SlideShare URL columns from DataFrame."""
    col_name = custom_name_col
    col_url = custom_url_col

    # 1. Use user-specified columns if provided and valid
    if col_name and col_name not in df.columns:
        console.print(f"[yellow]Warning: Specified name column '{col_name}' not found. Auto-detecting...[/yellow]")
        col_name = None
    if col_url and col_url not in df.columns:
        console.print(f"[yellow]Warning: Specified url column '{col_url}' not found. Auto-detecting...[/yellow]")
        col_url = None

    # 2. Match by header names
    for col in df.columns:
        c_str = str(col).strip().lower()
        if not col_name and any(k in c_str for k in ['company', 'name', 'title', 'firm', 'startup', 'organization']):
            col_name = col
        if not col_url and any(k in c_str for k in ['link', 'url', 'slideshare', 'ppt', 'deck', 'presentation', 'href']):
            col_url = col

    # 3. If URL column not detected, inspect values for slideshare or http
    if not col_url:
        for col in df.columns:
            sample_vals = df[col].dropna().astype(str).head(10).tolist()
            if any('slideshare' in val.lower() or 'http' in val.lower() for val in sample_vals):
                col_url = col
                break

    # 4. If name column not detected, pick the first column that isn't the URL column
    if not col_name:
        for col in df.columns:
            if col != col_url:
                col_name = col
                break

    return col_name, col_url


def fetch_slideshow_metadata(page, url: str):
    """
    Load SlideShare page with Playwright and extract slide information.
    Handles Next.js __NEXT_DATA__ and DOM fallbacks.
    """
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(1.5)  # Give JS time to hydrate
    except Exception as e:
        console.print(f"   [yellow]Page load error: {e}[/yellow]")
        return None

    # 1. Try extracting from __NEXT_DATA__
    try:
        data = page.evaluate("""
            () => {
                const el = document.getElementById('__NEXT_DATA__');
                return el ? JSON.parse(el.textContent) : null;
            }
        """)
        if data:
            slideshow = data.get("props", {}).get("pageProps", {}).get("slideshow")
            if slideshow:
                slides_info = slideshow.get("slides", {})
                total_slides = slideshow.get("totalSlides", 0)
                image_location = slides_info.get("imageLocation")
                slide_title = slides_info.get("title")
                host = slides_info.get("host", "https://image.slidesharecdn.com")

                if total_slides and image_location and slide_title:
                    return {
                        "type": "template",
                        "host": host,
                        "imageLocation": image_location,
                        "title": slide_title,
                        "totalSlides": total_slides,
                        "deckTitle": slideshow.get("title", slide_title)
                    }
    except Exception as e:
        console.print(f"   [yellow]Could not read __NEXT_DATA__: {e}[/yellow]")

    # 2. Fallback: Query slide image elements directly from the DOM
    try:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight / 3);")
        time.sleep(1)
        page.evaluate("window.scrollTo(0, 0);")

        img_urls = page.evaluate("""
            () => {
                const urls = [];
                const images = document.querySelectorAll('img');
                for (const img of images) {
                    const src = img.currentSrc || img.src || img.getAttribute('data-src') || '';
                    if (src.includes('slidesharecdn.com') && src.includes('-')) {
                        urls.push(src);
                    }
                }
                return urls;
            }
        """)
        if img_urls:
            return {
                "type": "direct_urls",
                "urls": img_urls,
                "deckTitle": page.title()
            }
    except Exception as e:
        console.print(f"   [yellow]DOM fallback error: {e}[/yellow]")

    return None


def download_single_slide(session, host, image_location, title, slide_num):
    """Attempt downloading a single slide image across available resolutions."""
    for q, w in RESOLUTIONS:
        img_url = f"{host}/{image_location}/{q}/{title}-{slide_num}-{w}.jpg"
        try:
            resp = session.get(img_url, timeout=12)
            if resp.status_code == 200 and len(resp.content) > 1000:
                return slide_num, resp.content
        except Exception:
            continue
    return slide_num, None


def download_deck_as_pdf(page, url: str, output_path: Path, max_workers: int = 8) -> tuple[bool, int, str]:
    """
    Downloads a SlideShare presentation and compiles it into a PDF.
    Returns: (success: bool, slide_count: int, error_message: str)
    """
    try:
        meta = fetch_slideshow_metadata(page, url)
        if not meta:
            return False, 0, "Failed to retrieve presentation metadata"

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        })

        slide_images = {}
        total_slides = 0

        if meta["type"] == "template":
            total_slides = meta["totalSlides"]
            host = meta["host"]
            image_loc = meta["imageLocation"]
            slide_title = meta["title"]

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [
                    executor.submit(download_single_slide, session, host, image_loc, slide_title, i)
                    for i in range(1, total_slides + 1)
                ]
                for f in as_completed(futures):
                    s_num, data = f.result()
                    if data:
                        slide_images[s_num] = data

        elif meta["type"] == "direct_urls":
            urls = meta["urls"]
            total_slides = len(urls)

            def fetch_url(idx, u):
                try:
                    r = session.get(u, timeout=12)
                    if r.status_code == 200:
                        return idx, r.content
                except Exception:
                    pass
                return idx, None

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(fetch_url, i, u) for i, u in enumerate(urls, 1)]
                for f in as_completed(futures):
                    s_num, data = f.result()
                    if data:
                        slide_images[s_num] = data

        if not slide_images:
            return False, 0, "No slide images could be downloaded"

        # Compile all ordered images to PDF using Pillow
        pil_images = []
        for i in range(1, total_slides + 1):
            if i in slide_images:
                try:
                    img = Image.open(io.BytesIO(slide_images[i])).convert("RGB")
                    pil_images.append(img)
                except Exception as e:
                    console.print(f"   [yellow]Warning: Could not process slide {i}: {e}[/yellow]")

        if not pil_images:
            return False, 0, "Failed to assemble slide images into PDF"

        output_path.parent.mkdir(parents=True, exist_ok=True)
        pil_images[0].save(
            str(output_path),
            save_all=True,
            append_images=pil_images[1:],
            quality=95
        )

        return True, len(pil_images), ""

    except Exception as e:
        return False, 0, str(e)


def main():
    parser = argparse.ArgumentParser(description="Download SlideShare pitch decks to PDF based on Excel sheet.")
    parser.add_argument("--file", "-f", type=str, default=None, help="Path to input Excel (.xlsx, .xls) or CSV file.")
    parser.add_argument("--output", "-o", type=str, default="downloads", help="Output directory for downloaded PDFs.")
    parser.add_argument("--name-col", type=str, default=None, help="Column name containing Company Name.")
    parser.add_argument("--url-col", type=str, default=None, help="Column name containing SlideShare URL.")
    parser.add_argument("--limit", "-l", type=int, default=None, help="Limit number of decks to download (for testing).")
    parser.add_argument("--start", "-s", type=int, default=1, help="Start from row index (1-based, default: 1).")
    parser.add_argument("--force", action="store_true", help="Force re-download even if PDF already exists.")
    parser.add_argument("--headless", action="store_true", default=True, help="Run browser in headless mode.")
    args = parser.parse_args()

    console.rule("[bold cyan]SlideShare Pitch Deck PDF Downloader[/bold cyan]")

    # 1. Locate input file
    input_file = args.file
    if not input_file:
        # Search current working directory for spreadsheet files
        candidates = list(Path(".").glob("*.xlsx")) + list(Path(".").glob("*.xls")) + list(Path(".").glob("*.csv"))
        candidates = [c for c in candidates if not c.name.startswith("download_report") and not c.name.startswith("~")]
        if candidates:
            input_file = str(candidates[0])
            console.print(f"[green]Auto-detected input file:[/green] {input_file}")
        else:
            console.print("[red]Error: No Excel (.xlsx, .xls) or CSV file found in current directory.[/red]")
            console.print("Please place your Excel file here or pass it via: [cyan]python download_pitchdecks.py --file your_file.xlsx[/cyan]")
            sys.exit(1)

    # 2. Read the file
    input_path = Path(input_file)
    if not input_path.exists():
        console.print(f"[red]Error: File '{input_path}' not found.[/red]")
        sys.exit(1)

    try:
        if input_path.suffix.lower() == ".csv":
            df = pd.read_csv(input_path)
        else:
            df = pd.read_excel(input_path)
    except Exception as e:
        console.print(f"[red]Error reading '{input_path}': {e}[/red]")
        sys.exit(1)

    # Clean column headers (strip whitespace and UTF-8 BOM)
    df.columns = [str(c).strip().replace('\ufeff', '') for c in df.columns]

    total_rows = len(df)
    console.print(f"Loaded [bold]{total_rows}[/bold] rows from [cyan]{input_path.name}[/cyan]")

    # 3. Detect columns
    col_name, col_url = find_columns(df, args.name_col, args.url_col)
    if not col_name or not col_url:
        console.print("[red]Error: Could not identify Company Name and/or SlideShare Link columns.[/red]")
        console.print(f"Available columns: {list(df.columns)}")
        console.print("Please specify them using: [cyan]--name-col \"ColName\" --url-col \"ColURL\"[/cyan]")
        sys.exit(1)

    console.print(f"Using Company Column: [bold green]'{col_name}'[/bold green]")
    console.print(f"Using SlideShare Link Column: [bold green]'{col_url}'[/bold green]")

    # Apply --start and --limit slicing
    start_offset = max(1, args.start) - 1
    df_to_process = df.iloc[start_offset:]
    if args.limit:
        df_to_process = df_to_process.head(args.limit)

    console.print(f"Processing [bold]{len(df_to_process)}[/bold] rows (from row {start_offset + 1} to {start_offset + len(df_to_process)})\n")

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 4. Process each row with Playwright
    results = []
    
    with sync_playwright() as p:
        console.print("[cyan]Launching Chromium browser...[/cyan]")
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        for idx, row in df_to_process.iterrows():
            raw_name = row[col_name]
            raw_url = row[col_url]

            # Validation
            if pd.isna(raw_name) or not str(raw_name).strip():
                company_name = f"Company_Row_{idx + 1}"
            else:
                company_name = str(raw_name).strip()

            if pd.isna(raw_url) or not str(raw_url).strip() or "slideshare" not in str(raw_url).lower():
                console.print(f"[{idx+1}/{len(df)}] [yellow]Skipping '{company_name}': Invalid or missing SlideShare URL ({raw_url})[/yellow]")
                results.append({
                    "Index": idx + 1,
                    "Company": company_name,
                    "URL": str(raw_url) if not pd.isna(raw_url) else "",
                    "Status": "Skipped (Invalid URL)",
                    "Slides": 0,
                    "Output File": ""
                })
                continue

            target_filename = sanitize_filename(company_name) + ".pdf"
            target_path = output_dir / target_filename

            if target_path.exists() and not args.force and target_path.stat().st_size > 10240:
                console.print(f"[{idx+1}/{len(df)}] [green]Exists:[/green] '{target_filename}' already downloaded. Skipping.")
                results.append({
                    "Index": idx + 1,
                    "Company": company_name,
                    "URL": str(raw_url),
                    "Status": "Already Exists",
                    "Slides": "-",
                    "Output File": str(target_path)
                })
                continue

            console.print(f"[{idx+1}/{len(df)}] [bold blue]Processing:[/bold blue] '{company_name}' -> {raw_url}")

            start_t = time.time()
            success, slides_count, err = download_deck_as_pdf(page, str(raw_url).strip(), target_path)
            duration = time.time() - start_t

            if success:
                console.print(f"      [bold green][SUCCESS][/bold green] Downloaded {slides_count} slides -> [cyan]{target_path}[/cyan] ({duration:.1f}s)")
                results.append({
                    "Index": idx + 1,
                    "Company": company_name,
                    "URL": str(raw_url),
                    "Status": "Success",
                    "Slides": slides_count,
                    "Output File": str(target_path)
                })
            else:
                console.print(f"      [bold red][FAILED][/bold red] {err}")
                results.append({
                    "Index": idx + 1,
                    "Company": company_name,
                    "URL": str(raw_url),
                    "Status": f"Failed: {err}",
                    "Slides": 0,
                    "Output File": ""
                })

            # Check page health and recreate if needed
            try:
                page.evaluate("() => 1")
            except Exception:
                try:
                    page.close()
                except Exception:
                    pass
                page = context.new_page()

            # Save report incrementally every 10 processed items
            if len(results) > 0 and len(results) % 10 == 0:
                try:
                    pd.DataFrame(results).to_excel(output_dir / "download_report.xlsx", index=False)
                except Exception:
                    pass

        browser.close()

    # 5. Output Summary Table & Report
    console.print("\n")
    table = Table(title="Download Summary Report", show_header=True, header_style="bold magenta")
    table.add_column("#", style="dim", width=4)
    table.add_column("Company", style="bold")
    table.add_column("Slides", justify="right")
    table.add_column("Status")
    table.add_column("Output File")

    for r in results:
        status_color = "green" if r["Status"] == "Success" else ("yellow" if "Exists" in r["Status"] else "red")
        table.add_row(
            str(r["Index"]),
            r["Company"],
            str(r["Slides"]),
            f"[{status_color}]{r['Status']}[/{status_color}]",
            r["Output File"]
        )

    console.print(table)

    # Save summary report to Excel
    report_df = pd.DataFrame(results)
    report_path = output_dir / "download_report.xlsx"
    try:
        report_df.to_excel(report_path, index=False)
        console.print(f"\n[green]Summary report saved to:[/green] [bold]{report_path}[/bold]")
    except Exception:
        pass


if __name__ == "__main__":
    main()
