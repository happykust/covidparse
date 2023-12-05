import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Here we will save our results grouped by region to future saving to excel document
results = {}

# Open urls.txt file with urls of news
with open("urls.txt", "r") as f:
    urls = list(map(lambda x: x.replace("\n", ""), f.readlines()))

for url in urls:
    print(f"Garbage data from {url} -> ", end="")

    # Get the HTML document of news
    response = requests.get(url)
    # Parsing HTML document
    parsed_html = BeautifulSoup(response.text, "lxml")

    try:
        # Title with period of news
        title = parsed_html.find("div", "article-detail__content").find("h3").text
        # HTML table body
        table = parsed_html.find("div", "u-table-cv__wrapper").find("tbody")

        # Date of the end period
        end_period_date = title.split("-")[1].replace(")", "").replace("\n", "").replace(" ", "")

        # Go through all regions (rows) in table and garbage data
        for el in table.find_all("tr")[1:]:
            # All values for current region
            values = el.find_all("td")

            region = values[0].text.strip()
            hospitalized = values[1].text.strip().replace(" ", "")
            recovered = values[2].text.strip().replace(" ", "")
            revealed = values[3].text.strip().replace(" ", "")
            died = values[4].text.strip().replace(" ", "")

            # Add region to dict
            if region not in results:
                results[region] = {}

            # Save data to dict
            results[region][end_period_date] = {
                "hospitalized": hospitalized,
                "recovered": recovered,
                "revealed": revealed,
                "died": died
            }
    except Exception as e:
        print(f"Error: {e}")
        continue

    print("OK")

# Create new Excel Workbook
wb = Workbook()

# Grab the active worksheet
ws = wb.active

# Write headers
ws.append(["Дата окончания периода", "Регион", "Госпитализировано", "Выздоровело", "Выявлено", "Умерло"])

# Write data from dict
for region in results:
    if "Наименование субъекта" in region:
        continue
    for end_period_date in results[region]:
        data = results[region][end_period_date]
        ws.append([end_period_date, region, data["hospitalized"], data["recovered"],
                   data["revealed"], data["died"]])

# Save workboot to file
wb.save("result.xlsx")
