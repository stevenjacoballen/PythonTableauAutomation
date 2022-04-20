from datetime import datetime  # get datetime object
import requests  # web page request
from bs4 import BeautifulSoup  # web scraping
import re  # regular expressions
import openpyxl  # Excel file manipulation
import pyautogui  # GUI interaction (mouse movements and clicks)


# Get date & time
now = datetime.now()


# Scrape DarkSky.net for current temperature
data = requests.get('https://darksky.net/forecast/40.5155,-112.033/us12/en')
soup = BeautifulSoup(data.text, 'html.parser')
scraped_temperature = soup.select('span .summary.swap')


# Extract just the current temperature
temperature_regex = re.compile(r'\d\d')
object_match = temperature_regex.search(scraped_temperature[0].text)
current_temperature = object_match.group()


# Append timestamp and temperature to Tableau's data source (Excel).
workbook = openpyxl.load_workbook('/Users/steven/Desktop/tableau_weather.xlsx')
page = workbook.active
new_data = [[now, int(current_temperature)]]
for info in new_data:
    page.append(info)
workbook.save('/Users/steven/Desktop/tableau_weather.xlsx')
workbook.close()


# Update and publish Tableau workbook to Tableau Public.
# Pause for 1 second after Excel upload, and activate emergency fail-safe (throw curser to upper left-corner of screen
# to stop script)
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

# Open Tableau
pyautogui.moveTo(1320, 1439, duration=1)
pyautogui.PAUSE = 1
pyautogui.moveTo(1320, 1398, duration=.5)
pyautogui.click(1320, 1398)
pyautogui.PAUSE = 10

# Open Tableau workbook
pyautogui.moveTo(411, 290, duration=1)
pyautogui.PAUSE = 1
pyautogui.click(411, 290)
pyautogui.PAUSE = 2

# Click 'Data Source' button
pyautogui.moveTo(78, 1401, duration=1)
pyautogui.click(78, 1401)

# Click 'refresh' button
pyautogui.moveTo(201, 69, duration=1)
pyautogui.click(201, 69)

# Click 'Automated Dashboard' tab
pyautogui.moveTo(148, 1401, duration=1)
pyautogui.click(148, 1401)

# Click 'File'
pyautogui.moveTo(181, 12, duration=1)
pyautogui.click(181, 12)

# Click 'Save to Tableau Public'
pyautogui.moveTo(191, 124, duration=1)
pyautogui.click(191, 124)
pyautogui.PAUSE = 15

# Close web browser
pyautogui.moveTo(19, 44, duration=1)
pyautogui.click(19, 44)
pyautogui.PAUSE = 1

# Close Tableau
pyautogui.moveTo(12, 33, duration=1)
pyautogui.click(12, 33)
