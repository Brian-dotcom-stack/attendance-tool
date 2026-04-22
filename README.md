# Attendance Tracker Automation Tool

Automates the process of updating an **Excel Attendance Master Tracker** by syncing leave/cancellation data from:

- WhatsApp messages or text files  
- Sage HR API  

It parses leave records, matches employees using fuzzy name matching, and updates the correct monthly sheet cells with proper absence codes.

---

## Features

- 📄 Import leave data from WhatsApp text or Sage HR
- 🤖 Optional AI-powered parsing (Claude API)
- 🔍 Fuzzy name matching (handles nicknames and typos)
- 📊 Updates Excel monthly sheets (Jan–Dec structure)
- 🧠 Auto-detects leave types (Sick, Holiday, Emergency, etc.)
- 📝 Adds remarks for special cases (e.g. part-time workers)
- 🧪 Dry-run mode (preview without writing)

---

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
````

---

### 2. Configure settings

```bash
cp config.json.example config.json
```

Then edit:

```json
{
  "anthropic_api_key": "",
  "sage_hr_api_key": "",
  "sage_hr_subdomain": "",
  "fuzzy_match_threshold": 80,
  "default_status_code": "UA"
}
```

---

### 3. Run the tool

#### From WhatsApp file

```bash
python main.py --source whatsapp --input cancellations.txt --excel tracker.xlsx
```

#### From raw WhatsApp text

```bash
python main.py --source whatsapp --text "Winifred cancelled from 24 March to 30 March" --excel tracker.xlsx
```

#### From Sage HR (current month)

```bash
python main.py --source sageHR --excel tracker.xlsx
```

#### From Sage HR (specific month)

```bash
python main.py --source sageHR --excel tracker.xlsx --month 3 --year 2026
```

#### Dry run (no Excel changes)

```bash
python main.py --source whatsapp --input cancellations.txt --excel tracker.xlsx --dry-run
```

#### Output to new file

```bash
python main.py --source whatsapp --input cancellations.txt --excel tracker.xlsx --output tracker_updated.xlsx
```

---

## Excel Status Codes

| Code | Meaning                  |
| ---- | ------------------------ |
| P    | Present                  |
| UA   | Unauthorised Absence     |
| AA   | Authorised Absence       |
| SA   | Sickness Absence         |
| EA   | Emergency Absence        |
| OR   | Other Reason             |
| WO   | Weekend Off / Rest Day   |
| H    | Holiday / Approved Leave |

---

## How it works

1. Parses cancellation/leave records
2. Matches employee names using fuzzy matching
3. Locates correct month sheet (Jan–Dec)
4. Updates correct day columns in Excel
5. Writes status codes (UA, AA, SA, etc.)
6. Adds remarks for special cases (e.g. part-time staff)

---

## Special Rules

### Part-time (2 days/week)

If detected:

* Marks full leave range as `UA`
* Adds remark:

  ```
  Working 2 days/week — update P days manually
  ```

---

## Configuration Options

| Key                     | Description                                 |
| ----------------------- | ------------------------------------------- |
| `anthropic_api_key`     | Claude API key for advanced message parsing |
| `sage_hr_api_key`       | Sage HR API access key                      |
| `sage_hr_subdomain`     | Your Sage HR company subdomain              |
| `fuzzy_match_threshold` | Name matching sensitivity (default: 80)     |
| `default_status_code`   | Default Excel status (usually UA)           |

---

## Getting API Keys

### Claude (Anthropic)

1. Visit [https://console.anthropic.com](https://console.anthropic.com)
2. Create an account
3. Generate API key
4. Add to `config.json`

> Optional — tool works without it using regex parsing.

---

### Sage HR API

1. Login to Sage HR
2. Go to **Settings → Integrations → API**
3. Enable API access
4. Copy API key + subdomain

---

## Troubleshooting

### Unmatched employees

Causes:

* Nickname mismatch (e.g. "Favour" vs full name)
* Spelling differences
* Lower fuzzy threshold (`70–75` recommended)

---

### Sage HR returns no data

Check:

* API key validity
* Subdomain correctness
* API permissions enabled

Test:

```bash
curl -H "X-Auth-Token: YOUR_KEY" https://yourcompany.sage.hr/api/v1/employees
```

---

## Project Structure

```
attendance_tool/
│── main.py
│── excel_updater.py
│── sage_hr.py
│── whatsapp_parser.py
│── config.json
│── requirements.txt
│── tracker.xlsx
```

---

## Notes

* Designed for Excel attendance systems with monthly sheets
* Uses fuzzy matching to reduce manual corrections
* Built for HR/admin automation workflows

---



