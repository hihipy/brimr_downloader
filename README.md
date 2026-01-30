# BRIMR Downloader

If you've ever spent time manually downloading Excel files from the BRIMR website one by one, you know how tedious it can be. This Python script automates that work so you can get straight to analyzing NIH funding data.

What would take hours of clicking becomes a single organized download.

------

## The Challenge

The [Blue Ridge Institute for Medical Research (BRIMR)](https://brimr.org/) provides invaluable NIH funding rankings, but accessing the data comes with some friction:

- **Manual Downloads:** Each Excel file must be downloaded individually, and there are 70+ files per year across nearly two decades of data.
- **No Organization:** Files download with inconsistent names and no folder structure, leaving you with a mess in your Downloads folder.
- **Time-Consuming:** Downloading just five years of data manually could take an hour or more of repetitive clicking.
- **Easy to Miss Files:** With so many links on each page, it's easy to accidentally skip files or download duplicates.

------

## The Solution

This script uses browser automation to visit each year's BRIMR page and systematically download every Excel file, organizing them into a clean folder structure.

### Why Browser Automation?

Selenium-based browser automation ensures reliable, complete downloads.

- **Handles Dynamic Content:** BRIMR pages load content dynamically, which simple HTTP requests can miss. Browser automation sees exactly what you would see.
- **Reliable Downloads:** By automating a real browser, downloads work exactly as they would if you clicked each link yourself.
- **User-Friendly GUI:** The script uses a familiar desktop interface. No command-line interaction required. Just check the years you want and click Download.

------

## Features

- **Simple GUI:** A clean desktop interface lets you select years and monitor progress.
- **Single Chrome Instance:** Reuses one browser session for all downloads, making multi-year downloads faster.
- **Smart Organization:** Automatically categorizes files into logical folders (School Rankings, Clinical Departments, PI Rankings, etc.).
- **Skip Existing Files:** Won't re-download files you already have. Perfect for updating your dataset with new years.
- **Cancel Anytime:** Stop mid-download without losing what's already been saved.
- **Logging:** Every download is logged to a timestamped file for debugging.

------

## Getting Started

### You'll Need Python

Install **Python 3.10** or newer from the official [Python website](https://www.python.org/).

### Running the Downloader

1. **Set Up Your Environment:**

   Open your terminal or command prompt, navigate to the project folder, and run:

   **macOS / Linux:**

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install selenium webdriver-manager requests
   ```

   **Windows:**

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install selenium webdriver-manager requests
   ```

2. **Run the Script:**

   ```bash
   python brimr_downloader.py
   ```

3. **Select Your Years:**

   The GUI will open and automatically detect available years from the BRIMR website. Check the years you want (or use "Select All" / "Recent 5 Years").

4. **Download:**

   Hit the Download button and watch the progress bar. Your files will be organized and ready when it's done.

------

## Output Structure

The downloader creates an organized folder structure in your Downloads folder.

| Folder / File                                  | Description                                                                        |
| ---------------------------------------------- | ---------------------------------------------------------------------------------- |
| `BRIMR_Data/`                                  | Root folder containing all downloaded files, organized by year.                    |
| `BRIMR_Data/2024/`                             | Each selected year gets its own folder.                                            |
| `BRIMR_Data/2024/01_Source_Data/`              | Worldwide rankings, contracts, and comprehensive institutional data.               |
| `BRIMR_Data/2024/02_School_Rankings/`          | Medical schools, nursing, pharmacy, dentistry, and other school types.             |
| `BRIMR_Data/2024/03_Department_Summaries/`     | Aggregated views of funding by department.                                         |
| `BRIMR_Data/2024/04_Basic_Science/`            | Biochemistry, genetics, neurosciences, pharmacology, and more.                     |
| `BRIMR_Data/2024/05_Clinical_Depts/`           | Medicine, surgery, pediatrics, psychiatry, and 20+ specialties.                    |
| `BRIMR_Data/2024/06_PI_Rankings/`              | Individual researcher rankings by department and institution.                      |
| `BRIMR_Data/2024/07_Geographic/`               | Rankings by state, city, and institution location.                                 |
| `BRIMR_Data/2024/08_Other/`                    | Top Ten lists, COVID awards, MERIT awards, and other datasets.                     |
| `brimr_downloader_*.log`                       | Timestamped log of every file downloaded.                                          |

------

## Options

| Option             | Description                                                      |
| ------------------ | ---------------------------------------------------------------- |
| **Output Folder**  | Change where files are saved (defaults to Downloads/BRIMR_Data). |
| **Headless Mode**  | Run Chrome invisibly in the background (enabled by default).     |
| **Year Selection** | Check individual years or use quick-select buttons.              |

------

## Troubleshooting

**Chrome doesn't start / ChromeDriver error**
The script automatically downloads the correct ChromeDriver for your Chrome version. Make sure Google Chrome is installed and your internet connection is working.

**Some years show no files**
BRIMR data has a natural lag. The most recent year's data may not be available yet. The script will report "Years without data: X" if any selected years are empty.

**Download seems stuck**
Some files are larger and take longer. Check the log file for progress. You can also disable headless mode to watch the browser work.

------

## Requirements

- Python 3.10+
- Google Chrome
- Internet connection

**Python Packages:**
- `selenium` - Browser automation
- `webdriver-manager` - Automatic ChromeDriver management
- `requests` - HTTP requests for year detection

------

## License

This project is licensed under [CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/).

You are free to:
- Use, share, and adapt this work
- Use it at your job

Under these terms:
- **Attribution** — Credit the original author
- **NonCommercial** — No selling or commercial products
- **ShareAlike** — Derivatives must use the same license
