# 📊 Google Sheets ETL & State Automator

A custom-built Google Apps Script (GAS) suite designed to parse, deduplicate, and visually synchronize large datasets without hitting Google's API execution limits. 

This repository also includes a custom Python CLI tool (`generate_data.py`) utilizing the `Faker` library to generate tens of thousands of rows of realistic, "messy" analytics data for end-to-end load testing.

## 🚀 The Problem & The Solution
**The Problem:** Processing weekly web analytics data exports often requires hours of manual deduplication, URL sanitization, and category-splitting. Furthermore, when data is split into multiple sub-tabs, maintaining a "single source of truth" for visual highlights and status tracking becomes impossible.

**The Solution:** This tool adds a native UI menu to Google Workspace that executes a full ETL (Extract, Transform, Load) pipeline entirely within the browser, reducing a multi-hour workflow to 3 clicks.

## 🧠 Core Engineering Skills Demonstrated
* **Performance Optimization:** Utilizes in-memory array batching (`getValues()` / `setValues()`) rather than iterative cell-by-cell loops to bypass GAS API execution limits and achieve `O(N)` time complexity.
* **State Management & Synchronization:** Features bidirectional syncing algorithms that map Hex color states across disparate spreadsheet tabs, allowing for absolute or additive state mirroring.
* **Data Cleaning (Regex):** Sanitizes and normalizes erroneous URL strings (e.g., stripping duplicate `.html.html` extensions) during the deduplication phase.
* **Testing Infrastructure:** Built a standalone Python CLI tool with `argparse` to seed massive, controlled datasets for stress-testing the JavaScript pipeline.

## 🛠️ Toolkit Features

1. **✨ Data Prep & Deduplication:** Cleans erroneous paths and aggregates matching user metrics while preserving the formatting of the highest-volume row.
2. **🎨 Visual Syntax Highlighting:** Parses URL strings and dynamically applies localized RichText styles (Hex color mapping) to distinguish folders, slashes, and file extensions.
3. **📂 Automated Category Splitting:** Reads the Master dataset into memory, groups by category, expands sheet dimensions dynamically, and stamps out categorized sub-tabs while retaining perfect Master column widths.
4. **🔄 Bidirectional State Syncing:** * `Push`: Forces categorized sub-tabs to inherit color states from the Master database.
    * `Pull Highlights`: Harvests new status highlights from sub-tabs back to the Master.
    * `Absolute Sync`: Mirrors the exact current state (including color removals) from sub-tabs to Master.
5. **📝 Review Extraction:** Scans an entire dataset and extracts only rows containing active color states into a centralized review pipeline.


## 💻 Setup & Testing

**1. Prepare Your Google Sheet (Quick Start)**
1. Create a new Google Sheet.
2. **CRITICAL:** Rename the primary tab to exactly `MASTER` (all caps). The script relies on this precise naming convention to execute its parsing and extraction functions.
3. Import the provided `tests/sample_data.csv` into the `MASTER` tab.

**2. Deploy the Apps Script**
1. Navigate to `Extensions` > `Apps Script`.
2. Paste the contents of `src/Code.js` into the editor and click Save.
3. Select the `onOpen` function from the top dropdown and click **Run** to authenticate the script.
4. Return to your spreadsheet to access the new **🚀 Developer Tools** menu.

**3. Generate Heavy Load Data (Optional)**
If you want to test the in-memory array batching performance, use the Python CLI tool to generate a larger dataset. Ensure you have the `faker` library installed (`pip install faker`).
```bash
# Generate a standard 500-row test dataset
python tests/generate_data.py --rows 500 --output test_data.csv

# Generate a 10,000-row dataset for performance stress-testing
python tests/generate_data.py --rows 10000 --output heavy_test_data.csv
