# BRIMR Downloader



Hello there! If you've ever spent time manually downloading Excel files from the BRIMR website one by one, you know how tedious it can be. This automated Python utility was built to take that repetitive work off your plate so you can get straight to analyzing NIH funding data.

This script transforms what would be hours of clicking into a single, organized download with just a few clicks.

------



## The Challenge: Why Is Downloading BRIMR Data So Tedious?



The Blue Ridge Institute for Medical Research (BRIMR) provides invaluable NIH funding rankings, but accessing the data comes with some friction:

- **Manual Downloads:** Each Excel file must be downloaded individually—and there are 70+ files per year across nearly two decades of data.
- **No Organization:** Files download with inconsistent names and no folder structure, leaving you with a jumbled mess in your Downloads folder.
- **Time-Consuming:** Downloading just five years of data manually could take an hour or more of repetitive clicking.
- **Easy to Miss Files:** With so many links on each page, it's easy to accidentally skip files or download duplicates.

This manual process is not only time-consuming but also prone to human error, potentially leaving gaps in your dataset.

------



## The Solution: Your Automated Data Collector



This Python script provides a robust, one-click solution to these challenges. It uses browser automation to visit each year's BRIMR page and systematically download every Excel file, organizing them into a clean folder structure.

The script handles all the tedious "data collection" work, allowing you to move directly from deciding what years you need to having analysis-ready files.



### **Why Browser Automation?**



The choice of Selenium-based browser automation was deliberate to ensure reliable, complete downloads.

- **Handles Dynamic Content:** BRIMR pages load content dynamically, which simple HTTP requests can miss. Browser automation sees exactly what you would see.
- **Reliable Downloads:** By automating a real browser, downloads work exactly as they would if you clicked each link yourself.
- **User-Friendly GUI:** The script uses a familiar desktop interface—no command-line interaction required. Just check the years you want and click Download.

------



## Key Features & Benefits



- ✨ **User-Friendly GUI:** A clean desktop interface lets you select years and monitor progress visually.
- 🚀 **Single Chrome Instance:** Reuses one browser session for all downloads, making multi-year downloads significantly faster.
- 📂 **Smart Organization:** Automatically categorizes files into logical folders (School Rankings, Clinical Departments, PI Rankings, etc.).
- 🔄 **Skip Existing Files:** Won't re-download files you already have, perfect for updating your dataset with new years.
- ❌ **Cancel Anytime:** A cancel button lets you stop mid-download without losing what's already been saved.
- 📋 **Complete Logging:** Every download is logged to a timestamped file for full transparency and debugging.

------



## Getting Started: Your 5-Minute Guide



### **First, You'll Need Python**



If you don't already have it, you'll need to install **Python 3.10** or newer. You can download it from the official [Python website](https://www.python.org/).



### **Now, Let's Run the Downloader**



1. **Prepare Your Environment:**

   Open your terminal or command prompt, navigate to the project folder, and run the following commands to create a virtual environment and install the required packages.

   **On macOS / Linux:**

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install selenium webdriver-manager requests
   ```

   **On Windows:**

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install selenium webdriver-manager requests
   ```

2. **Run the Downloader:**

   With your virtual environment active, run the script from your terminal:

   ```bash
   python brimr_downloader.py
   ```

3. **Select Your Years:**

   The GUI will open and automatically detect available years from the BRIMR website. Check the years you want to download (or use "Select All" / "Recent 5 Years").

4. **Click Download and Relax!**

   Hit the Download button and watch the progress bar. Your files will be organized and ready when it's done.

------



## The Final Product: What's in the Box? 📦



The downloader produces a complete, organized folder structure in your Downloads folder.

| Folder / File                              | What It Is & Why You Need It                                                                 |
| ------------------------------------------ | -------------------------------------------------------------------------------------------- |
| **`BRIMR_Data/`**                          | **Your Data Home.** The root folder containing all downloaded files, organized by year.      |
| **`BRIMR_Data/2024/`**                     | **Year Folders.** Each selected year gets its own folder for easy navigation.                |
| **`BRIMR_Data/2024/01_Source_Data/`**      | **Raw Data.** Worldwide rankings, contracts, and comprehensive institutional data.           |
| **`BRIMR_Data/2024/02_School_Rankings/`**  | **School Rankings.** Medical schools, nursing, pharmacy, dentistry, and other school types.  |
| **`BRIMR_Data/2024/03_Department_Summaries/`** | **Department Rollups.** Aggregated views of funding by department.                       |
| **`BRIMR_Data/2024/04_Basic_Science/`**    | **Basic Science.** Biochemistry, genetics, neurosciences, pharmacology, and more.            |
| **`BRIMR_Data/2024/05_Clinical_Depts/`**   | **Clinical Departments.** Medicine, surgery, pediatrics, psychiatry, and 20+ specialties.    |
| **`BRIMR_Data/2024/06_PI_Rankings/`**      | **Principal Investigators.** Individual researcher rankings by department and institution.   |
| **`BRIMR_Data/2024/07_Geographic/`**       | **Geographic Data.** Rankings by state, city, and institution location.                      |
| **`BRIMR_Data/2024/08_Other/`**            | **Special Files.** Top Ten lists, COVID awards, MERIT awards, and other unique datasets.     |
| **`brimr_downloader_*.log`**               | **Your Download Receipt.** A timestamped log of every file downloaded for full transparency. |

------



## Configuration Options



The GUI provides several options to customize your download:

| Option | Description |
| ------ | ----------- |
| **Output Folder** | Change where files are saved (defaults to Downloads/BRIMR_Data). |
| **Headless Mode** | Run Chrome invisibly in the background (enabled by default). |
| **Year Selection** | Check individual years or use quick-select buttons for common ranges. |

------



## Troubleshooting



**Chrome doesn't start / ChromeDriver error**
The script automatically downloads the correct ChromeDriver for your Chrome version. Make sure you have Google Chrome installed and your internet connection is working.

**Some years show no files**
BRIMR data has a natural lag—the most recent year's data may not be available yet. The script will report "Years without data: X" if any selected years have no files.

**Download seems stuck**
Some files are larger and take longer to download. Check the log file for detailed progress. You can also try disabling headless mode to watch the browser work.

------



## Requirements



- Python 3.10+
- Google Chrome (installed on your system)
- Internet connection

**Python Packages:**
- `selenium` — Browser automation
- `webdriver-manager` — Automatic ChromeDriver management
- `requests` — HTTP requests for year detection

------



## License



BRIMR Downloader © 2026 – Distributed under the [Creative Commons Attribution 4.0 International License](https://creativecommons.org/licenses/by/4.0/).