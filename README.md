# Attendance Tracker Automation Tool

Automates the process of updating an **Excel Attendance Master Tracker** by syncing leave and cancellation data from:

- WhatsApp messages or text files  
- Sage HR API  

It parses leave records, matches employees using fuzzy name matching, and updates the correct monthly sheet cells with appropriate absence codes.

---

## 🚀 Features

- 📄 Import leave data from WhatsApp text or Sage HR  
- 🤖 Optional AI-powered parsing (Claude API)  
- 🔍 Fuzzy name matching (handles nicknames, partial names, typos)  
- 📊 Updates Excel monthly sheets (Jan–Dec structure)  
- 🧠 Auto-detects leave types (Sick, Holiday, Emergency, etc.)  
- 📝 Adds remarks for special cases (e.g. part-time workers)  
- 🧪 Dry-run mode (preview changes without writing to Excel)  

---

## ⚡ Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
````

---

### 2. Configure settings

Create a file called `config.json` in the root directory:

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

#### Output to a new file

```bash
python main.py --source whatsapp --input cancellations.txt --excel tracker.xlsx --output tracker_output.xlsx
```

---

## 📊 Example Output

```
📲  Parsing WhatsApp cancellation messages…
✅  Found 5 cancellation record(s):

📝  Writing to Excel: tracker_output.xlsx

✅  Done!
   Cells updated  : 12
   Unmatched names: 0
```

---

## 📁 Project Structure

```
attendance-tool/
│── main.py
│── excel_updater.py
│── attendance_tool/
│   ├── parsers/
│   │   ├── whatsapp_parser.py
│   │   ├── sage_hr.py
│   ├── config.py
│── requirements.txt
│── cancellations.txt (example)
│── README.md
```

---

## 📌 Excel Template Requirements

This tool expects an Excel file with:

* Monthly sheets named: **Jan → Dec**
* Employee names in **Column B**
* Day columns starting from **Column D**
* A structured attendance format matching the template

---

## 🧾 Excel Status Codes

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

## ⚙️ Configuration Options

| Key                   | Description                             |
| --------------------- | --------------------------------------- |
| anthropic_api_key     | Claude API key for advanced parsing     |
| sage_hr_api_key       | Sage HR API access key                  |
| sage_hr_subdomain     | Your Sage HR company subdomain          |
| fuzzy_match_threshold | Name matching sensitivity (default: 80) |
| default_status_code   | Default Excel status (usually UA)       |

---

## 🔌 API Setup

### Claude (Anthropic)

1. Visit [https://console.anthropic.com](https://console.anthropic.com)
2. Create an account
3. Generate an API key
4. Add it to `config.json`

> Optional — tool works without it using regex parsing.

---

### Sage HR API

1. Log in to Sage HR
2. Go to **Settings → Integrations → API**
3. Enable API access
4. Copy API key and subdomain

---

## ⚠️ Limitations

* Requires a predefined Excel template structure
* Fuzzy matching may occasionally match similar names incorrectly
* Does not currently detect duplicate or overlapping leave entries

---

## 🛠 Troubleshooting

### Unmatched employees

Possible causes:

* Nicknames vs full names
* Spelling differences
* Low fuzzy match threshold

Try lowering threshold to **70–75**.

---

### Sage HR returns no data

Check:

* API key validity
* Subdomain correctness
* API permissions

Test with:

```bash
curl -H "X-Auth-Token: YOUR_KEY" https://yourcompany.sage.hr/api/v1/employees
```

---

## 🧠 How It Works

1. Parses cancellation/leave records
2. Matches employee names using fuzzy matching
3. Locates correct month sheet (Jan–Dec)
4. Updates correct day columns in Excel
5. Writes absence codes (UA, AA, SA, etc.)
6. Adds remarks for special cases

---

## 📌 Notes

* Built for HR/admin automation workflows
* Reduces manual Excel updates significantly
* Designed for real-world messy data (partial names, typos, etc.)

---

