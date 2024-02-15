import win32com.client
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime

url = "https://www.xtb.com/en/market-analysis/market-calendar"

driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(10)

page_source = driver.page_source
driver.quit()

soup = BeautifulSoup(page_source, 'html.parser')
rows = soup.select('ul.market-calendar-list li.js-row')

data = []

for row in rows:
    impact = row.select_one('div.col-impact div')['class'][0].replace('impact-', '')

    if impact == '3':
        time = row.select_one('div.col-time').text.strip()
        impact = row.select_one('div.col-impact div')['class'][0].replace('impact-', '')
        country_code = row.select_one('div.col-country div.flag-icon')['title']
        indicator = row.select_one('div.col-indicator').text.strip()
        period = row.select_one('div.col-detail.d-none.d-md-block.col-3').text.strip()
        previous_value = row.select_one('div.col-previous').text.strip().split()[1]
        current_value = row.select_one('div.col-current').text.strip().split()[1]
        forecast_elem = row.select_one('div.col-forecast')
        forecast_value = forecast_elem.text.strip().split()[1] if forecast_elem else None

        data.append([time, impact, country_code, indicator, period, previous_value, current_value, forecast_value])

table_html = "<table border='1' style='width:100%;'>"
table_html += "<tr style='background-color: #79A4BF; text-align: center;'>"
table_html += "<th>Time</th><th>Impact</th><th>Country Code</th><th>Indicator</th><th>Period</th><th>Previous Value</th><th>Current Value</th><th>Forecast Value</th>"
table_html += "</tr>"

for row in data:
    color = '#BDD2DF'
    table_html += f"<tr style='background-color: {color}; text-align: center;'>" + "".join(f"<td>{value}</td>" for value in row) + "</tr>"

table_html += "</table>"

intro_text = "Messieurs,<br><br>Vous trouverez ci-dessous le calendrier économique du jour :<br><br>"

current_date = datetime.now().strftime("%d/%m/%Y")

subject = f"Calendrier Économique du {current_date}"

email_body = f"<html><body>{intro_text}{table_html}</body></html>"

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

drafts_folder = namespace.GetDefaultFolder(16)

mail_item = outlook.CreateItem(0)
mail_item.Subject = subject
mail_item.TO = ""
mail_item.CC = ""
mail_item.HTMLBody = email_body
mail_item.Save()

mail_item.Display()

print("Brouillon créé avec succès !")
