# 📊 Product Adoption Data Scraper

A scalable automation pipeline to extract, process, and store product adoption data from WebEngage dashboards across multiple regions and channels.

---

## 🚀 Overview

This project automates:

* 🔐 Login via SSO (Selenium-based session handling)
* 🔑 Access management (auto request-access handling)
* 📡 API data extraction (Overview + Channel-level)
* 📅 Monthly backfill (Oct’25 → Feb’26)
* 🧹 Data parsing & transformation
* 📊 Google Sheets integration (region + month-wise tabs)
* 🔁 Resume capability (no reprocessing)

---

## 🏗️ Project Structure

```
Product_Adaption_Data/
│
├── auth/
│   ├── login.py              # Selenium login
│   ├── cookies.py            # Extract session cookies
│
├── access/
│   ├── request_access.py     # Access API handler
│
├── data/
│   ├── fetch_overview.py
│   ├── fetch_channel_campaign.py
│   ├── parser.py
│   ├── parser_channels.py
│   ├── load_lc.py
│   ├── sheet_writer.py
│
├── config/
│   ├── settings.py           # URLs, months, configs
│   ├── channel_config.py     # Channel endpoints
│
├── utils/
│   ├── logger.py
│   ├── date_filter.py
│
├── jobs/
│   ├── run_scrapper.py       # Main execution file
│
├── progress_channels.json    # Resume tracking
└── README.md
```

---

## ⚙️ Features

### ✅ Smart Access Handling

* Detects `403` errors
* Automatically requests access
* Retries API call

---

### ✅ Session Recovery

* Detects expired cookies
* Prompts re-login once
* Continues execution (no restart needed)

---

### ✅ Early Stop Optimization

* Stops pagination once data is older than required month
* Prevents unnecessary API calls

---

### ✅ Resume Capability

* Tracks completed License Codes
* Skips already processed entries
* Saves progress in:

  ```
  progress_channels.json
  ```

---

### ✅ Google Sheets Output

Data is stored as:

```
Sheet Tabs:
Oct'25 GLOBAL
Nov'25 GLOBAL
Dec'25 GLOBAL
Jan'26 GLOBAL
Feb'26 GLOBAL
```

Each row includes:

* License Code
* Channel
* Campaign ID
* Title
* Status
* Metrics (Sent, Delivered, Clicks, etc.)
* Month

---

## 🌍 Supported Regions

* INDIA
* GLOBAL
* KSA

Run per region:

```bash
python3 jobs/run_scrapper.py --region GLOBAL
```

---

## 📡 Channels Supported

* PUSH_NOTIFICATION
* SMS
* EMAIL
* WEB_PUSH
* WHATSAPP
* FACEBOOK

---

## 🧠 How It Works

### Step 1: Login

* Opens browser
* User completes SSO login
* Session cookies captured

---

### Step 2: Data Extraction

For each:

```
Region → Channel → License Code → Pages
```

* Fetch campaigns
* Handle pagination
* Apply early stop logic

---

### Step 3: Data Filtering

* Uses `createdOn`
* Filters campaigns within month range

---

### Step 4: Storage

* Parsed data pushed to Google Sheets
* Batched for performance

---

## 🛡️ Rate Limiting & Safety

To avoid blocking:

* Random delays between API calls
* Batch cooldown every 10 LCs
* Pagination limits (`max 50 pages`)
* Retry logic instead of spamming

---

## ⚠️ Known Limitations

* Requires manual SSO login (once per session)
* Cookie expires after some time
* Channel APIs do not support direct date filters
  → Workaround: filter using `createdOn`

---

## 🔧 Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

---

### 2. Google Sheets Setup

* Create Service Account
* Download JSON credentials
* Share sheet access with service email

---

### 3. Configure

Update:

```
config/settings.py
```

* Base URLs
* Role IDs
* Backfill months

---

## ▶️ Running the Project

### Run for a region:

```bash
python3 jobs/run_scrapper.py --region GLOBAL
```

---

## 📈 Future Improvements

* 🔄 Fully automated login (no manual input)
* ⚡ Parallel processing (multi-threading)
* 🧠 Smart retry queue for failed LCs
* 📊 Unified dashboard (overview + channels)
* ⏱️ Scheduler (cron-based automation)

---

## 👨‍💻 Author

Hitesh Kotian

---

## ⭐ Notes

This project demonstrates:

* API reverse engineering
* Automation design
* Scalable data pipelines
* Real-world analytics workflow

---
