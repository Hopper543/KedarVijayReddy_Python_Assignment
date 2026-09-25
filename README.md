# Lead Gen Script

Pulls developer/company leads from GitHub (public API), cleans them up and saves to Excel.

## Note

I used the GitHub REST API (search + user profiles) as the data source since it's public and has the fields we need: name, email, website, location, company. The data is loaded into pandas and cleaned up. That means trimming spaces, fixing URLs, checking emails, finding emails and LinkedIn links written in bios, removing duplicates, and filling empty fields with "N/A". The results go into a formatted Excel file with a Leads sheet and a Summary sheet. It can also push to a Google Sheet.
Extras: if a lead has their own domain, the script guesses possible emails (first@, first.last@, flast@). It also has a `--every` flag to rerun on a schedule.

Tools: Python, requests, pandas, openpyxl, gspread (optional).

## Setup

```
pip install -r requirements.txt
```

Optional but recommended: set a GitHub token to get past the 60 requests/hr limit.

```
set GH_TOKEN=your_token
```

## Usage

```
python leads.py                                   # 30 leads, default query
python leads.py -q "location:bangalore type:org" -n 50
python leads.py --every 24                        # rerun every 24h
python leads.py --sheet "Leads"                   # also push to google sheets
```

Any [GitHub user search query](https://docs.github.com/en/search-github/searching-on-github/searching-users) works with `-q`. For example, `type:org` gets companies and orgs instead of people.

Output:
- `output/leads.xlsx` has the cleaned leads plus a summary sheet
- `output/raw.csv` has the raw data before cleaning

### Google Sheets

Create a service account in Google Cloud and enable the Sheets + Drive API. Download its key as `creds.json` into this folder, then run with `--sheet`. If this isn't set up, that step is skipped and the Excel file is still saved.

### Scheduling on Windows

Instead of keeping `--every` running, you can use Task Scheduler:

```
schtasks /create /tn leads /tr "python E:\Jarurat\leads.py" /sc daily /st 09:00
```
