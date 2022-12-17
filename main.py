import requests
from bs4 import BeautifulSoup
from openpyxl import *


def scrap_event_name(event_name):
    eventNames = []
    for i in range(1, len(event_name), 2):
        name = event_name[i].text
        name = name.replace("\n\n", '')
        name = name.replace("\n", '')
        name = name.replace("              ", '')
        eventNames.append(name)
    return eventNames


def scrap_event_description(event_description):
    eventDescriptions = []
    for i in range(0, len(event_description) - 2):
        name = event_description[i].text
        name = name.replace("\n              ", '')
        name = name.replace("\r\n\n            ", '')
        name = name.replace(" \n            ", '')
        eventDescriptions.append(name)
    return eventDescriptions


def scrap_event_date(event_date):
    eventDates = []
    for i in range(len(event_date)):
        eventDates.append(eventDate[i].text)
    return eventDates


def scrap_event_time(event_time):
    eventTimes = []
    for i in range(len(event_time)):
        eventTimes.append(event_time[i].text)
    return eventTimes


result = requests.get("https://www.flexjobs.com/events?time=upcoming&topics=29")
pageSrc = result.content
soup = BeautifulSoup(pageSrc, "html.parser")

eventName = soup.find_all("a", {"class": "stretched-link"})  # return list
# print(eventName)
eventDate = soup.find_all("span", {"data-time-format": "EEE, MMM d, yyyy"})
# print(eventDate)
eventStartTime = soup.find_all("span", {"data-time-format": "h:mm a"})
# print(eventStartTime)
eventEndTime = soup.find_all("span", {"data-time-format": "h:mm a ZZZZ"})
# print(eventEndTime)
eventDescription = soup.find_all("p", {"class": "m-0"})
# print(eventDescription)

nameOfEvents = scrap_event_name(eventName)
dateOfEvents = scrap_event_date(eventDate)
startTimeOfEvents = scrap_event_time(eventStartTime)
endTimeOfEvents = scrap_event_time(eventEndTime)
descriptionOfEvents = scrap_event_description(eventDescription)

wb = Workbook()
wb = load_workbook("CareerGapEvents.xlsx")
sheet = wb.active
for i in range(len(nameOfEvents)):
    cell = 'A' + str(i + 1)
    sheet[cell] = nameOfEvents[i]

for i in range(len(dateOfEvents)):
    cell = 'B' + str(i + 1)
    sheet[cell] = dateOfEvents[i]

for i in range(len(startTimeOfEvents)):
    cell = 'C' + str(i + 1)
    sheet[cell] = startTimeOfEvents[i]

for i in range(len(endTimeOfEvents)):
    cell = 'D' + str(i + 1)
    sheet[cell] = endTimeOfEvents[i]

for i in range(len(descriptionOfEvents)):
    cell = 'E' + str(i + 1)
    sheet[cell] = descriptionOfEvents[i]

wb.save("CareerGapEvents.xlsx")
