import asyncio
import httpx
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule


API_URL = "https://pfebooks.com/wp-json/wp/v2/posts"
OUTPUT_FILE = "entreprises_pfe.xlsx"
HEADERS = {"User-Agent": "Mozilla/5.0"}

# -----------------------------------------------------
# SCRAPER
# -----------------------------------------------------
async def fetch_page(client, page):
    """Fetches a single page of posts from the API."""
    try:
        r = await client.get(API_URL, params={"per_page": 100, "page": page}, timeout=30)
        r.raise_for_status()
        return r.json()
    except httpx.RequestError as e:
        print(f"An error occurred while requesting page {page}: {e}")
        return []
    except Exception as e:
        print(f"An unexpected error occurred on page {page}: {e}")
        return []

async def scrape_all():
    """Scrapes all posts and returns them as a list of dictionaries."""
    results = []
    async with httpx.AsyncClient(headers=HEADERS) as client:
        page = 1
        print("Starting scraping...")
        while True:
            data = await fetch_page(client, page)
            if not data:
                break
            for item in data:
                title = item.get("title", {}).get("rendered", "No Title")
                results.append({
                    "Name": title,
                    "Project Submitted": "",
                    "Response": "",
                    "Statut": "Non fait"
                })
            print(f"Page {page} scraped successfully.")
            page += 1
    print(f"Scraping finished. Found {len(results)} items.")
    return results

# -----------------------------------------------------
# EXCEL BUILDER
# -----------------------------------------------------
def save_excel(items):
    """Saves the list of items to a formatted Excel file."""
    if not items:
        print("No items to save. Excel file not generated.")
        return
    df = pd.DataFrame(items)
    df.to_excel(OUTPUT_FILE, index=False)
    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active
    RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    dv = DataValidation(type="list", formula1='"Non fait,Fait"')
    ws.add_data_validation(dv)
    dv.add("D2:D" + str(ws.max_row))
    formatting_range = "A2:D" + str(ws.max_row)
    ws.conditional_formatting.add(
        formatting_range,
        FormulaRule(formula=['$D2="Fait"'], fill=GREEN_FILL)
    )
    ws.conditional_formatting.add(
        formatting_range,
        FormulaRule(formula=['$D2="Non fait"'], fill=RED_FILL)
    )
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 15
    wb.save(OUTPUT_FILE)
    print(f"✔ Excel generated successfully: {OUTPUT_FILE}")

# -----------------------------------------------------
# MAIN
# -----------------------------------------------------
async def main():
    """Main function to run the scraper and Excel builder."""
    items = await scrape_all()
    save_excel(items)

if __name__ == "__main__":
    asyncio.run(main())
