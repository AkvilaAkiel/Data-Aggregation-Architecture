# OTT Data Aggregation Architecture 📺⚙️

> **Note:** *This is a sanitized version of the script created for portfolio showcase. All internal Spreadsheet IDs, API keys, and proprietary endpoints have been removed to comply with NDA.*

## 📌 Overview
In a large-scale OTT streaming environment, managing live sports broadcasting schedules across multiple regions involves massive amounts of fragmented data. 

**The Problem:** Manual collection of match data (CDN IDs, internal IDs, start times, commentator assignments) from multiple scattered Google Sheets was consuming **~40 man-hours per week** and was highly prone to human error, leading to potential stream misconfigurations.

**The Solution:** A custom, highly optimized **Google Apps Script (GAS)** that acts as a Single Source of Truth (SSOT). It dynamically scans multiple data sources, normalizes unstructured data, performs flexible cross-sheet matching, and aggregates everything into a standardized, "ready-to-use" dashboard for the operations team.

## 🚀 Business Impact
* **Efficiency:** Completely eliminated manual data entry, saving **~40 hours per week** of operations workload.
* **Reliability:** Reduced human error to zero in schedule mapping, preventing broadcasting delays and wrong stream routing.
* **Scalability:** The architecture effortlessly handles hundreds of concurrent events across 15+ geographic regions.

## ✨ Key Features
* **🧠 Smart Header Mapping:** Automatically identifies and maps columns regardless of language (UA/RU/EN) or minor naming variations (e.g., "Start Time", "Час", "Начало").
* **⏳ Dynamic Date-Windowing:** Automatically filters and processes only relevant events (Current Day + 7-day outlook) to reduce API payload and execution time.
* **🔗 Relational Data Matching:** Acts as an advanced programmatic `VLOOKUP`. It matches events from raw provider schedules with internal broadcasting sheets using dynamically generated `Auto IDs`.
* **🧹 Data Normalization:** Built-in parsers sanitize complex time/date strings, fix nomenclature typos (e.g., auto-converting `MF10` to `F10`), and append regional Geo-tags based on sheet metadata.
* **⚡ Optimized Execution:** Minimizes Google Sheets API calls by utilizing bulk `getDisplayValues()` reads, in-memory Map/Set data structures, and targeted loop breaks.

## 🛠 Tech Stack
* **Language:** JavaScript (ES6+ features)
* **Environment:** Google Apps Script (GAS)
* **APIs:** Google Sheets API, V8 Runtime

## 📋 How It Works (Logic Flow)
1. **Phase 1 (Fetch & Parse):** Scans all active sheets in the source document, ignores hidden sheets, identifies headers via fuzzy matching, and extracts matches falling within the 7-day window.
2. **Phase 2 (Normalization):** Generates unique `Auto IDs` for relational mapping based on stream endpoints.
3. **Phase 3 (Cross-Matching):** Connects to the external "Schedule" database, maps commentators, engines, and channel names to the raw events.
4. **Phase 4 (Render):** Clears the destination board and renders a perfectly formatted, aggregated daily schedule, automatically inserting visual breaks between different days.
