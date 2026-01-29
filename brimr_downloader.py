#!/usr/bin/env python3
"""
BRIMR NIH Funding Data Downloader

Downloads Excel files from BRIMR using browser automation.
Automatically detects available years from the website.

Features:
    - Thread-safe UI updates
    - Single Chrome instance (faster multi-year downloads)
    - Logging to file
    - Configurable timeouts
    - Cancel button
    - Cross-platform (Windows, macOS, Linux)

Requirements:
    pip install selenium webdriver-manager requests

Usage:
    python brimr_downloader.py
"""

from __future__ import annotations

import logging
import platform
import queue
import re
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Callable
from urllib.parse import unquote, urlparse

import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

if TYPE_CHECKING:
    from tkinter import Event

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# =============================================================================
# Configuration Constants - Adjust these if needed
# =============================================================================

PAGE_LOAD_TIMEOUT: int = 15
DOWNLOAD_TIMEOUT: int = 90
EXTRA_RENDER_WAIT: float = 2.0
DOWNLOAD_CHECK_INTERVAL: float = 0.3
UI_QUEUE_POLL_MS: int = 100

BASE_URL_TEMPLATE: str = "https://brimr.org/brimr-rankings-of-nih-funding-in-{year}/"

FILE_CATEGORIES: dict[str, list[str]] = {
    # Source/master data files - comprehensive institutional data
    "01_Source_Data": [
        "worldwide", "worldwidebrimr", "brimrworldwide",
        "allorgs", "medicalschoolsonly", "medicalschools",
        "contracts", "somcontracts", "worldwidecontractsonly",
    ],
    # School-level rankings
    "02_School_Rankings": [
        "schoolofmedicine",
        "schoolofdentistry", "dentistry",
        "schoolofnursing", "nursing",
        "schoolofpublichealth",
        "schoolofpharmacy", "pharmacy",
        "schoolofveterinarymedicine", "schoolofverterinarymedicine",
        "schoolsofveterinarymedicine", "veterinarymedicine",
        "schoolofosteopathicmedicine",
        "schoolofalliedhealth",
        "hospitals",
        "otherhealthprofessions",
    ],
    # Department summary/rollup files
    "03_Department_Summaries": [
        "bydepartment", "bydepartmentr",
        "awardsbydepartment",
        "medicalschoolsandtheirdepartments",
        "medicalschoolsanddept", "medicalschoolanddept",
        "mastertemplatenihawards", "mastertemplatenihawardsr",
    ],
    # Basic science departments
    "04_Basic_Science": [
        "anatomycellbiol", "anatomycellbiology", "anatomycellbiologyr",
        "biochemistry", "biochemistryr",
        "biomedicalengineering",
        "genetics", "geneticsr",
        "microbiology", "microbiologyr",
        "neurosciences", "neurosciencesr",
        "pharmacology", "pharmacologyr",
        "physiology", "physiologyr",
        "otherbasicsciences",
    ],
    # Clinical departments
    "05_Clinical_Depts": [
        "anesthesiology", "anesthesiologyr",
        "dermatology", "dermatologyr",
        "emergencymedicine", "emergencymediciner",
        "familymedicine", "familymediciner",
        "medicine", "mediciner",
        "neurology", "neurologyr", "neurologyxls",
        "neurosurgery", "neurosurgeryr",
        "nutrition",
        "obgyn", "obgynr",
        "obstetrics", "obstetricsandgynecology", "obstetricsgynecology",
        "ophthalmology", "ophthalmologyr",
        "orthopedics", "orthopedicsr",
        "ent", "otolaryngology", "otolaryngologyr",
        "pathology", "pathologyr",
        "pediatrics", "pediatricsr",
        "physicalme", "physicalmed", "physicalmedicine", "physicalmediciner",
        "psychiatry", "psychiatryr",
        "publichealth", "publichealthr",
        "radiology", "radiologyr",
        "surgery", "surgeryr",
        "urology", "urologyr",
        "otherclinicalsciences",
    ],
    # Principal Investigator rankings
    "06_PI_Rankings": [
        "pi", "allpis", "pisbyrank", "principalinvestigator",
        "allorgdeptpi", "deptorgpi", "deptschoolpi", "deptschoolpir",
        "schooldeptpi",
        "contractspi",
        # Basic science PI files
        "anatomycellbiolpi", "anatomycellbiologypi",
        "biochemistrypi",
        "biomedicalengineeringpi",
        "geneticspi",
        "microbiologypi",
        "neurosciencespi",
        "pharmacologypi", "pharmacologypir",
        "physiologypi", "physiologypi2xls", "physiologypir",
        # Clinical PI files
        "anesthesiologypi",
        "dermatologypi",
        "emergencymedicinepi",
        "familymedicinepi",
        "medicinepi",
        "neurologypi",
        "neurosurgerypi", "neurosurgerypir",
        "obgynpi", "obgynpir",
        "obstetricspi", "obstetricsandgynecologypi", "obstetricsgynecologypi",
        "ophthalmologypi", "ophthalmologypir",
        "orthopedicspi", "orthopedicspir",
        "otolaryngologypi", "otolaryngologypir",
        "pathologypi", "pathologypir",
        "pediatricspi", "pediatricspir",
        "physicalmedicinepi", "physicalmedicinepir",
        "psychiatrypi", "psychiatrypir",
        "publichealthpi", "publichealthpir",
        "radiologypi", "radiologypir",
        "surgerypi", "surgerypir", "surgerypibcorrecte",
        "urologypi", "urologypir",
    ],
    # Geographic analysis files
    "07_Geographic": [
        "state", "statesandcountries",
        "city", "allcities", "citiesbyrankr",
        "institution", "allinstitutions", "allinstitutionsr",
        "organization",
        "fundingrankbystate",
        "percapitafundingbystate", "percapitafundingrankbystate",
    ],
    # Special/other files
    "08_Other": [
        "topten", "toptenr",
        "covidawards", "covid",
        "merit", "nihmerit",
    ],
}


# =============================================================================
# Logging Setup
# =============================================================================


def setup_logging(log_dir: Path) -> logging.Logger:
    """Configure logging to file and console."""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"brimr_downloader_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger("brimr_downloader")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(console_handler)

    logger.info(f"Log file: {log_file}")
    return logger


# =============================================================================
# Cross-Platform Utility Functions
# =============================================================================


def get_downloads_folder() -> Path:
    """Get the user's Downloads folder path (cross-platform)."""
    system = platform.system()

    if system == "Windows":
        import winreg
        try:
            sub_key = (
                r"SOFTWARE\Microsoft\Windows\CurrentVersion"
                r"\Explorer\Shell Folders"
            )
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                downloads = winreg.QueryValueEx(
                    key, "{374DE290-123F-4565-9164-39C4925E467B}"
                )[0]
                return Path(downloads)
        except (OSError, FileNotFoundError):
            pass

    elif system == "Darwin":
        downloads = Path.home() / "Downloads"
        if downloads.exists():
            return downloads

    if system == "Linux":
        try:
            import subprocess
            result = subprocess.run(
                ["xdg-user-dir", "DOWNLOAD"],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0 and result.stdout.strip():
                return Path(result.stdout.strip())
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

    return Path.home() / "Downloads"


def detect_available_years(logger: logging.Logger | None = None) -> list[int]:
    """
    Detect available ranking years by checking which year URLs actually exist.

    Probes individual year URLs from current year back to 2006 to find all valid years.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        ),
    }

    years: set[int] = set()
    current_year = datetime.now().year

    if logger:
        logger.info(
            f"Probing BRIMR for available years ({current_year} down to 2006)..."
        )

    # Probe every year from current down to 2006
    for year in range(current_year, 2005, -1):
        url = BASE_URL_TEMPLATE.format(year=year)
        try:
            # Try HEAD request first (faster)
            response = requests.head(
                url, headers=headers, timeout=6, allow_redirects=True
            )

            if response.status_code == 200:
                years.add(year)
                if logger:
                    logger.debug(f"Year {year} exists (HEAD)")
            elif response.status_code in (403, 405):
                # Some servers reject HEAD, try GET with stream=True (minimal download)
                response = requests.get(
                    url, headers=headers, timeout=8, allow_redirects=True, stream=True
                )
                if response.status_code == 200:
                    years.add(year)
                    if logger:
                        logger.debug(f"Year {year} exists (GET fallback)")

        except requests.RequestException as e:
            if logger:
                logger.debug(f"Year {year} probe failed: {e}")
            continue

    if years:
        if logger:
            logger.info(f"Detected {len(years)} years: {min(years)}-{max(years)}")
        return sorted(years, reverse=True)

    # Fallback - known working range
    if logger:
        logger.warning("Could not detect years, using known range 2006-2024")
    return list(range(2024, 2005, -1))


def categorize_file(filename: str) -> str:
    """Determine the category folder for a file based on its name."""
    raw = filename.lower()
    norm = raw.replace("-", "").replace("_", "")

    # Check PI files first (they often contain department names too)
    if "pi" in norm:
        # Check patterns that need underscores/hyphens preserved
        if any(tag in raw for tag in ["_pi_", "_pi.", "pi_2", "contractspi"]):
            return "06_PI_Rankings"
        # Check normalized patterns
        if any(tag in norm for tag in ["allorgdeptpi", "deptschoolpi", "schooldeptpi"]):
            return "06_PI_Rankings"

    # Build list of (pattern, category) sorted by pattern length descending
    # This ensures longer/more specific patterns match first
    all_patterns: list[tuple[str, str]] = []
    for category, patterns in FILE_CATEGORIES.items():
        if category == "06_PI_Rankings":
            continue
        for pattern in patterns:
            clean_pattern = pattern.replace("-", "").replace("_", "")
            all_patterns.append((clean_pattern, category))

    # Sort by pattern length descending (longest first)
    all_patterns.sort(key=lambda x: len(x[0]), reverse=True)

    for pattern, category in all_patterns:
        if pattern in norm:
            return category

    return "09_Uncategorized"


def sanitize_filename(filename: str) -> str:
    """Sanitize filename by removing query strings and invalid characters."""
    # Remove query string
    if "?" in filename:
        filename = filename.split("?", 1)[0]
    # Replace characters invalid on Windows and generally unsafe
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        filename = filename.replace(ch, "_")
    # Trim whitespace and trailing dots (Windows quirk)
    return filename.strip().strip(".")


# =============================================================================
# Chrome Driver Management
# =============================================================================


def create_chrome_driver(
    download_dir: Path,
    headless: bool = False,
    logger: logging.Logger | None = None
) -> webdriver.Chrome:
    """Create a Chrome WebDriver configured for automatic file downloads."""
    options = Options()
    download_dir_str = str(download_dir.resolve())

    prefs = {
        "download.default_directory": download_dir_str,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1200,800")

    if headless:
        options.add_argument("--headless=new")

    if logger:
        logger.info("Installing/locating ChromeDriver...")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir_str}
    )

    if logger:
        logger.info(f"Chrome driver ready (headless={headless})")

    return driver


def set_download_directory(driver: webdriver.Chrome, directory: Path) -> None:
    """Change the download directory for an existing Chrome driver."""
    directory.mkdir(parents=True, exist_ok=True)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(directory.resolve())}
    )


def wait_for_download_complete(
    download_dir: Path,
    timeout: int = DOWNLOAD_TIMEOUT,
    cancel_event: threading.Event | None = None
) -> Path | None:
    """Wait for a download to complete and return the downloaded file path."""
    end_time = time.time() + timeout

    while time.time() < end_time:
        if cancel_event and cancel_event.is_set():
            return None

        try:
            files = list(download_dir.iterdir())
        except OSError:
            time.sleep(DOWNLOAD_CHECK_INTERVAL)
            continue

        temp_suffixes = (".crdownload", ".tmp", ".part")
        excel_suffixes = (".xls", ".xlsx")
        temp_files = [f for f in files if f.suffix.lower() in temp_suffixes]
        excel_files = [
            f for f in files
            if f.suffix.lower() in excel_suffixes and f.is_file()
        ]

        if excel_files and not temp_files:
            return max(excel_files, key=lambda f: f.stat().st_mtime)

        time.sleep(DOWNLOAD_CHECK_INTERVAL)

    return None


# =============================================================================
# GUI Application
# =============================================================================


class BRIMRDownloaderApp:
    """Thread-safe GUI application for downloading BRIMR NIH Funding Excel files."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("BRIMR NIH Funding Downloader")
        self.root.geometry("620x780")
        self.root.resizable(True, True)

        self.downloads_folder: Path = get_downloads_folder()
        self.output_folder: Path = self.downloads_folder / "BRIMR_Data"
        self.year_vars: dict[int, tk.BooleanVar] = {}
        self.available_years: list[int] = []
        self.is_downloading: bool = False
        self.headless_var: tk.BooleanVar = tk.BooleanVar(value=False)

        self.ui_queue: queue.Queue[tuple[Callable, tuple]] = queue.Queue()
        self.cancel_event: threading.Event = threading.Event()
        self.logger: logging.Logger | None = None

        self._setup_ui()
        self._start_ui_queue_polling()
        self._detect_years_async()

    # -------------------------------------------------------------------------
    # Thread-Safe UI Updates
    # -------------------------------------------------------------------------

    def _start_ui_queue_polling(self) -> None:
        self._drain_ui_queue()

    def _drain_ui_queue(self) -> None:
        try:
            while True:
                fn, args = self.ui_queue.get_nowait()
                fn(*args)
        except queue.Empty:
            pass
        self.root.after(UI_QUEUE_POLL_MS, self._drain_ui_queue)

    def _enqueue_ui(self, fn: Callable, *args) -> None:
        self.ui_queue.put((fn, args))

    def _update_status(self, message: str, detail: str = "") -> None:
        def apply():
            self.status_label.config(text=message)
            self.detail_label.config(text=detail)
        self._enqueue_ui(apply)

    def _update_progress(self, value: int, maximum: int) -> None:
        def apply():
            self.progress_bar["maximum"] = maximum
            self.progress_bar["value"] = value
        self._enqueue_ui(apply)

    def _show_info(self, title: str, message: str) -> None:
        self._enqueue_ui(lambda: messagebox.showinfo(title, message))

    def _show_error(self, title: str, message: str) -> None:
        self._enqueue_ui(lambda: messagebox.showerror(title, message))

    def _set_buttons_state(self, downloading: bool) -> None:
        def apply():
            self.download_btn.config(state="disabled" if downloading else "normal")
            self.cancel_btn.config(state="normal" if downloading else "disabled")
        self._enqueue_ui(apply)

    # -------------------------------------------------------------------------
    # UI Setup
    # -------------------------------------------------------------------------

    def _setup_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._create_header(main_frame)
        self._create_folder_info(main_frame)
        self._create_year_selection(main_frame)
        self._create_options(main_frame)
        self._create_progress_section(main_frame)
        self._create_buttons(main_frame)

    def _create_header(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent, text="BRIMR NIH Funding Data Downloader",
            font=("Segoe UI", 14, "bold"),
        ).pack(pady=(0, 5))

        ttk.Label(
            parent, text="Downloads Excel files using Chrome browser automation",
            font=("Segoe UI", 10),
        ).pack(pady=(0, 5))

        ttk.Label(
            parent, text=f"Platform: {platform.system()} {platform.machine()}",
            font=("Segoe UI", 8, "italic"), foreground="gray",
        ).pack(pady=(0, 10))

    def _create_folder_info(self, parent: ttk.Frame) -> None:
        folder_frame = ttk.Frame(parent)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            folder_frame, text="Save to:", font=("Segoe UI", 9)
        ).pack(side=tk.LEFT)

        self.folder_label = ttk.Label(
            folder_frame, text=str(self.output_folder),
            font=("Segoe UI", 9, "italic"), foreground="gray",
        )
        self.folder_label.pack(side=tk.LEFT, padx=(5, 10))

        ttk.Button(
            folder_frame, text="Change...",
            command=self._change_output_folder, width=10,
        ).pack(side=tk.LEFT)

    def _change_output_folder(self) -> None:
        folder = filedialog.askdirectory(
            initialdir=self.output_folder, title="Select Output Folder"
        )
        if folder:
            self.output_folder = Path(folder)
            self.folder_label.config(text=str(self.output_folder))

    def _create_year_selection(self, parent: ttk.Frame) -> None:
        self.years_frame = ttk.LabelFrame(
            parent, text="Select Years (detecting...)", padding="10"
        )
        self.years_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        btn_frame = ttk.Frame(self.years_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            btn_frame, text="Select All", command=self._select_all
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            btn_frame, text="Deselect All", command=self._deselect_all
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(
            btn_frame, text="Recent 5 Years", command=self._select_recent_5
        ).pack(side=tk.LEFT)

        self.canvas = tk.Canvas(self.years_frame, height=180)
        scrollbar = ttk.Scrollbar(
            self.years_frame, orient="vertical", command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._bind_mousewheel()

        self.loading_label = ttk.Label(
            self.scrollable_frame,
            text="Detecting available years from BRIMR website...",
            font=("Segoe UI", 9, "italic"),
        )
        self.loading_label.pack(pady=20)

    def _bind_mousewheel(self) -> None:
        system = platform.system()
        if system == "Linux":
            self.canvas.bind_all(
                "<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units")
            )
            self.canvas.bind_all(
                "<Button-5>", lambda e: self.canvas.yview_scroll(1, "units")
            )
        else:
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event: Event) -> None:
        if platform.system() == "Darwin":
            delta = -1 * event.delta
        else:
            delta = -1 * (event.delta // 120)
        self.canvas.yview_scroll(int(delta), "units")

    def _detect_years_async(self) -> None:
        def detect():
            years = detect_available_years()
            self._enqueue_ui(self._populate_years, years)
        threading.Thread(target=detect, daemon=True).start()

    def _populate_years(self, years: list[int]) -> None:
        self.available_years = years
        self.loading_label.destroy()

        year_range = f"{min(years)}-{max(years)}" if years else "None found"
        self.years_frame.configure(text=f"Select Years ({year_range} available)")

        for i, year in enumerate(years):
            var = tk.BooleanVar(value=False)
            self.year_vars[year] = var
            cb = ttk.Checkbutton(
                self.scrollable_frame, text=str(year), variable=var, width=8
            )
            cb.grid(row=i // 5, column=i % 5, sticky="w", padx=5, pady=2)

        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _create_options(self, parent: ttk.Frame) -> None:
        options_frame = ttk.LabelFrame(parent, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            options_frame,
            text="Run in headless mode (no visible browser window)",
            variable=self.headless_var,
        ).pack(anchor="w")

        ttk.Label(
            options_frame,
            text="Files organized: Year → Category → Files | Logs saved with downloads",
            font=("Segoe UI", 8, "italic"), foreground="gray",
        ).pack(anchor="w", pady=(5, 0))

    def _create_progress_section(self, parent: ttk.Frame) -> None:
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(progress_frame, text="Ready to download")
        self.status_label.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        self.detail_label = ttk.Label(progress_frame, text="", font=("Segoe UI", 8))
        self.detail_label.pack(anchor="w", pady=(5, 0))

    def _create_buttons(self, parent: ttk.Frame) -> None:
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 5))

        style = ttk.Style()
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        self.download_btn = ttk.Button(
            btn_frame, text="Download Selected Years",
            command=self._start_download, style="Accent.TButton",
        )
        self.download_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.cancel_btn = ttk.Button(
            btn_frame, text="Cancel",
            command=self._cancel_download, state="disabled",
        )
        self.cancel_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # -------------------------------------------------------------------------
    # Selection Helpers
    # -------------------------------------------------------------------------

    def _select_all(self) -> None:
        for var in self.year_vars.values():
            var.set(True)

    def _deselect_all(self) -> None:
        for var in self.year_vars.values():
            var.set(False)

    def _select_recent_5(self) -> None:
        self._deselect_all()
        for year in self.available_years[:5]:
            if year in self.year_vars:
                self.year_vars[year].set(True)

    def _get_selected_years(self) -> list[int]:
        return [year for year, var in self.year_vars.items() if var.get()]

    # -------------------------------------------------------------------------
    # Download Control
    # -------------------------------------------------------------------------

    def _start_download(self) -> None:
        selected_years = self._get_selected_years()

        if not selected_years:
            messagebox.showwarning("No Selection", "Please select at least one year.")
            return

        if self.is_downloading:
            return

        year_str = ", ".join(str(y) for y in sorted(selected_years, reverse=True))
        mode = "headless (no window)" if self.headless_var.get() else "visible"

        if not messagebox.askyesno(
            "Confirm Download",
            f"Download Excel files for: {year_str}?\n\n"
            f"Browser mode: {mode}\n"
            f"Files will be saved to:\n{self.output_folder}"
        ):
            return

        self.logger = setup_logging(self.output_folder)
        self.logger.info(f"Starting download for years: {year_str}")

        self.is_downloading = True
        self.cancel_event.clear()
        self._set_buttons_state(downloading=True)

        threading.Thread(
            target=self._download_years, args=(selected_years,), daemon=True
        ).start()

    def _cancel_download(self) -> None:
        if self.is_downloading:
            self.cancel_event.set()
            self._update_status("Cancelling...", "Waiting for current file to finish")
            if self.logger:
                self.logger.info("Cancel requested by user")

    def _download_years(self, years: list[int]) -> None:
        """Download Excel files for selected years using a single Chrome instance."""
        downloaded = 0
        skipped = 0
        failed = 0
        years_skipped = 0  # Years with no data or failed to load
        cancelled = False

        self.output_folder.mkdir(parents=True, exist_ok=True)
        temp_dir = self.output_folder / "_temp_downloads"
        temp_dir.mkdir(exist_ok=True)

        driver: webdriver.Chrome | None = None

        try:
            self._update_status("Starting Chrome browser...", "This may take a moment")
            driver = create_chrome_driver(
                temp_dir, headless=self.headless_var.get(), logger=self.logger
            )

            # Warmup: navigate to a blank page first to ensure driver is fully ready
            driver.get("about:blank")
            time.sleep(0.5)

            for year in sorted(years):  # Oldest to newest
                if self.cancel_event.is_set():
                    cancelled = True
                    break

                year_folder = self.output_folder / str(year)
                year_folder.mkdir(exist_ok=True)
                set_download_directory(driver, temp_dir)

                url = BASE_URL_TEMPLATE.format(year=year)
                self._update_status(f"Loading {year} page...", url)
                if self.logger:
                    self.logger.info(f"Processing year {year}")

                try:
                    driver.get(url)
                    WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    time.sleep(EXTRA_RENDER_WAIT)

                    # Wait a bit more for dynamic content to load
                    # Try to wait for at least one link to appear
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located(
                                (By.CSS_SELECTOR, "a[href*='.xls']")
                            )
                        )
                    except Exception:
                        pass  # Continue anyway, links might use different format

                except Exception as e:
                    if self.logger:
                        self.logger.error(f"Failed to load {year}: {e}")
                    self._update_status(f"⚠ {year}: Page failed to load", str(e)[:50])
                    years_skipped += 1
                    time.sleep(1)  # Brief pause so user sees the message
                    continue

                links = driver.find_elements(
                    By.CSS_SELECTOR,
                    "a[href$='.xls'], a[href$='.xlsx'], "
                    "a[href$='.XLS'], a[href$='.XLSX'], "
                    "a[href*='.xls?'], a[href*='.xlsx?']"
                )

                if self.logger:
                    self.logger.debug(
                        f"CSS selector found {len(links)} link elements for {year}"
                    )

                excel_urls: list[str] = []
                for link in links:
                    href = (link.get_attribute("href") or "").strip()
                    if any(ext in href.lower() for ext in [".xls", ".xlsx"]):
                        excel_urls.append(href)

                if self.logger:
                    self.logger.debug(
                        f"After filtering: {len(excel_urls)} Excel URLs for {year}"
                    )

                excel_urls = list(dict.fromkeys(excel_urls))

                if not excel_urls:
                    self._update_status(
                        f"⚠ {year}: No Excel files found", "Skipping this year"
                    )
                    if self.logger:
                        self.logger.warning(f"No Excel files found for {year}")
                    years_skipped += 1
                    time.sleep(1)  # Brief pause so user sees the message
                    continue

                self._update_status(
                    f"Found {len(excel_urls)} files for {year}", "Starting..."
                )
                if self.logger:
                    self.logger.info(f"Found {len(excel_urls)} files for {year}")

                for i, file_url in enumerate(excel_urls, 1):
                    if self.cancel_event.is_set():
                        cancelled = True
                        break

                    raw_name = unquote(Path(urlparse(file_url).path).name)
                    filename = sanitize_filename(raw_name)
                    category = categorize_file(filename)
                    category_folder = year_folder / category
                    final_path = category_folder / filename

                    if final_path.exists():
                        skipped += 1
                        self._update_status(
                            f"{year}: {i}/{len(excel_urls)}",
                            f"⏭ Exists: {category}/{filename}"
                        )
                        self._update_progress(i, len(excel_urls))
                        continue

                    self._update_status(
                        f"Downloading {year}: {i}/{len(excel_urls)}", f"→ {filename}"
                    )
                    self._update_progress(i, len(excel_urls))

                    try:
                        for f in temp_dir.iterdir():
                            try:
                                f.unlink()
                            except OSError:
                                pass

                        driver.get(file_url)
                        time.sleep(0.5)

                        downloaded_file = wait_for_download_complete(
                            temp_dir, cancel_event=self.cancel_event
                        )

                        if self.cancel_event.is_set():
                            cancelled = True
                            break

                        if downloaded_file and downloaded_file.exists():
                            category_folder.mkdir(exist_ok=True)
                            shutil.move(str(downloaded_file), str(final_path))
                            downloaded += 1
                            if self.logger:
                                self.logger.info(f"✓ {category}/{filename}")
                            self._update_status(
                                f"{year}: {i}/{len(excel_urls)}",
                                f"✓ {category}/{filename}"
                            )
                        else:
                            failed += 1
                            if self.logger:
                                self.logger.warning(f"⚠ Timeout: {filename}")

                    except Exception as e:
                        failed += 1
                        if self.logger:
                            self.logger.exception(f"Failed: {filename}")

                if cancelled:
                    break

        except Exception as e:
            if self.logger:
                self.logger.exception("Fatal error")
            self._show_error("Error", f"An error occurred:\n{e}")

        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

            shutil.rmtree(temp_dir, ignore_errors=True)
            self.is_downloading = False
            self._set_buttons_state(downloading=False)

            summary = (
                f"Downloaded: {downloaded} | Existed: {skipped} | "
                f"Failed: {failed} | Years without data: {years_skipped}"
            )

            if cancelled:
                self._update_status("⚠ Download cancelled", summary)
                self._show_info("Cancelled", f"Download cancelled.\n\n{summary}")
            else:
                self._update_status("✓ Download complete!", summary)
                self._update_progress(100, 100)
                self._show_info(
                    "Complete",
                    f"Downloaded {downloaded} files!\n"
                    f"Already existed: {skipped}\n"
                    f"Failed: {failed}\n"
                    f"Years without data: {years_skipped}\n\n"
                    f"Files saved to:\n{self.output_folder}"
                )

            if self.logger:
                self.logger.info(f"Finished. {summary}")


# =============================================================================
# Main Entry Point
# =============================================================================


def main() -> None:
    root = tk.Tk()
    root.geometry("620x780")

    if platform.system() == "Windows":
        try:
            root.iconbitmap(default="")
        except tk.TclError:
            pass

    BRIMRDownloaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()