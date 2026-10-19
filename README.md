# SlideShare Pitch Deck PDF Downloader

An automated Python script that reads an Excel sheet containing Company Names and SlideShare presentation URLs, downloads each presentation in high-definition (HD), compiles the slides into a PDF, and saves each file named after the company.

---

## Features

- **Automated Excel & CSV Support**: Automatically detects `.xlsx`, `.xls`, or `.csv` files.
- **Smart Column Detection**: Auto-detects columns for Company Name and SlideShare URL (or specify custom columns via CLI).
- **Safe Windows File Naming**: Automatically cleans illegal characters from company names to generate safe file names (`{Company_Name}.pdf`).
- **High-Definition Quality**: Downloads the highest available resolution for each slide (2048px / 1024px / 638px HD).
- **Fast Multi-Threaded Downloads**: Uses parallel thread workers to download slide images rapidly.
- **Duplicate Prevention**: Skips already downloaded decks to save time and bandwidth (override with `--force`).
- **Detailed Execution Report**: Outputs a formatted terminal table and exports a `download_report.xlsx` with download statuses and slide counts.

---

## Requirements & Setup

All required packages can be installed via:

```bash
pip install -r requirements.txt
playwright install chromium
```

---

## Usage

### 1. Simple Run (Auto-detects Excel file in folder)
Place your Excel file (e.g. `sample_pitchdecks.xlsx` or your own `.xlsx` file) in this directory and run:

```bash
python download_pitchdecks.py
```

### 2. Specify an Excel or CSV File
```bash
python download_pitchdecks.py --file "my_companies.xlsx"
```

### 3. Specify a Custom Output Directory
```bash
python download_pitchdecks.py --file "my_companies.xlsx" --output "pitch_decks"
```

### 4. Specify Custom Column Names
If your column headers differ from standard names:
```bash
python download_pitchdecks.py --name-col "Startup Name" --url-col "Pitch Deck Link"
```

### 5. Force Re-download
To overwrite already downloaded PDFs:
```bash
python download_pitchdecks.py --force
```

---

## Excel Sheet Format Example

| Company Name | SlideShare URL |
| :--- | :--- |
| Airbnb | `https://www.slideshare.net/slideshow/airbnb-first-pitch-deck-editable/45768374` |
| Sequoia Capital Pitch Deck | `https://www.slideshare.net/slideshow/sequoia-capital-pitchdecktemplate/46231251` |

---

## Output Structure

Downloaded files are saved in the `downloads/` directory:
```
downloads/
├── Airbnb.pdf
├── Sequoia_Capital_Pitch_Deck.pdf
└── download_report.xlsx
```
