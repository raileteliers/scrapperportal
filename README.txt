# Portal Inmobiliario Scraper (Selenium + Nordic JSON)

## Quick start

1) Install dependencies:

    pip install -r requirements.txt

2) Run:

    python searchdatahouse.py

## Change how many listings to collect

Edit in `searchdatahouse.py`:

    MAX_LISTINGS = 100

## Outputs

- Excel: `properties_las_condes_YYYY-MM-DD.xlsx`
- Checkpoint: `checkpoint_las_condes_YYYY-MM-DD.json`
- Debug HTML: `debug_selenium.html`

## Notes

- If the Excel file is open, the script writes to a new file with a timestamp suffix.
