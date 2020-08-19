############################ PACKAGE IMPORTS
import time
from selenium import webdriver
from bs4 import BeautifulSoup
import os
import datetime
import xlsxwriter
import gspread
from google.oauth2.service_account import Credentials

############################ GOOGLE INFO
# Gmail Email
global google_username
google_username = 'INSERT GMAIL EMAIL HERE'

# Gmail Password
global google_password
google_password = 'INSERT GMAIL PASSWORD HERE'

# Time to do Google Login & Authentication in seconds
# You might need to do some 2-factor authentication, so 15 seconds should be enough :)
global time_authentication
time_authentication = 15

# GOOGLE SHEET INFO
global google_sheet_workbook
google_sheet_workbook = "Medium Data Workbook"  # This is the name of your Google Sheets workbook that you have already created where you want to save the data
global google_sheet_workbook_sheet
google_sheet_workbook_sheet = "Data Sheet"  # This is the name of the Google Sheets sheet inside of your workbook

# Time to switch to new pages in seconds
global time_switchpages
time_switchpages = 10

global today
today = str(datetime.datetime.now().date())

############################ GOOGLE CHROMEDRIVE & API CREDS INFO
global path_chromedriver
path_chromedriver = "C:/Users/asus/Downloads/chromedriver_win32/chromedriver.exe"
# Download the chromedriver here: https://chromedriver.chromium.org/downloads
# Make sure you use the right one for your version of Chrome!

global google_api_creds_filepath
google_api_creds_filepath = "C:/Users/asus/PycharmProjects/MediumStats_Dashboarder/client_secret.json"
# To get your Google Sheets credentials .json file, follow the steps here: https://www.analyticsvidhya.com/blog/2020/07/read-and-update-google-spreadsheets-with-python/


############################ METHODS

def login_google():
    # File path of your chromedriver.exe file
    global path_chromedriver

    if len(path_chromedriver) == 0:
        print("ERROR!")
        print("Please download the Selenium Chromedriver and put the .exe filepath in the global variable spot above.")
        exit()

    if len(google_api_creds_filepath) == 0:
        print("ERROR!")
        print(
            "Please get a Google Sheets API Credentials file from here: https://cloud.google.com/docs/authentication/production#cloud-console")
        exit()

    # initialize Selenium webdriver
    global browser
    browser = webdriver.Chrome(path_chromedriver)

    # open new tab with link to sign-in options for Medium
    browser.get('https://medium.com/m/signin?redirect=https%3A%2F%2Fmedium.com%2F&operation=login')

    # this pauses/delays the code execution for () seconds;
    # very important - give time for page to load before searching out elements
    # adjust (increase/decrease) depending on internet speed?
    time.sleep(time_switchpages)

    # link to sign-in with Google - the 1st button on the page
    # Might not work. If not, you can manually login when the program asks you to do so.

    try:
        print("Trying to login to Medium with Google info...")
        # Will try to login to Medium.com using Google Login

        google_button = browser.find_element_by_xpath('//button[0]').click()

        time.sleep(time_switchpages)

        # enter your Google email
        browser.find_element_by_id('identifierId').send_keys(google_username)

        # click the "Next" button
        browser.find_element_by_id('identifierNext').click()

        time.sleep(time_switchpages + 10)

        # enter your account password
        browser.find_element_by_xpath("//input[@name='password'][@type='password']").send_keys(google_password)

        # sign-in to your Medium account with Google auth!
        browser.find_element_by_id('passwordNext').click()

        time.sleep(time_authentication)

    except:
        print("************************************")
        print("Automatic Google login didn't work.\nPlease login and press enter when you're on the Medium.com homepage.")

        enter_entry = "."

        while len(enter_entry) > 0:
            enter_entry = input("Waiting for you to press enter...")

    print("************************************")
    print("Let's start grabbing stats!")


#####
#####


def mkdir(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def clean_number(number):
    number = str(number).lower()

    try:

        if "fans" in number:
            number = number.replace("fans", "")

        if "fan" in number:
            number = number.replace("fan", "")

        if "k" in number:
            number = number.replace("k", "")
            number = float(number)
            number = number * 1000

        else:
            number = int(float(number))

    except:
        pass

    return number


def clean_time(time_text):
    time_text = str(time_text)
    has_hours = False
    has_min = False
    has_sec = False
    hours = 0
    minutes = 0
    seconds = 0

    if "hr" in time_text.lower():
        has_hours = True
        hr_position = time_text.find("hr")
        hours = int(time_text[0:hr_position - 1].lstrip().rstrip())

    if "min" in time_text.lower():
        has_min = True
        min_position = time_text.find("min")
        if has_hours == True:
            minutes = int(time_text[hr_position + 3:min_position - 1].lstrip().rstrip())
        else:
            minutes = int(time_text[0:min_position - 1].lstrip().rstrip())

    if "sec" in time_text.lower():
        sec_position = time_text.find("sec")
        if has_min == True:
            seconds = int(time_text[min_position + 3:sec_position - 1].lstrip().rstrip())
        else:
            seconds = int(time_text[0:sec_position - 1].lstrip().rstrip())

    # Returns Time in Hours
    return ((hours * 60) + minutes + (seconds / 60))


def write_to_excel(story_data):
    today = str(datetime.datetime.now().date())

    print(story_data)

    MasterSheet = "Medium Stats - " + str(today)

    print("*****************")
    print('Creating your Excel File...')
    parentDir = "Medium Stats Folder"
    mkdir(parentDir)
    fileName = "{parent}/{file}.xlsx".format(parent=parentDir, file=MasterSheet)

    workbook = xlsxwriter.Workbook(fileName)
    # ERROR!
    worksheet = workbook.add_worksheet()
    row = 1

    # Making Row 1 Headers
    worksheet.write('A' + str(row), 'Date of Data Pull')
    worksheet.write('B' + str(row), 'Title')
    worksheet.write('C' + str(row), 'Link')
    worksheet.write('D' + str(row), 'Date Published')
    worksheet.write('E' + str(row), 'Publication')
    worksheet.write('F' + str(row), 'Views')
    worksheet.write('G' + str(row), 'Earnings')
    worksheet.write('H' + str(row), 'Member Total Time Viewed (Minutes)')
    worksheet.write('I' + str(row), 'Average Time Viewed (Minutes)')
    worksheet.write('J' + str(row), 'Fans')

    row += 1

    for story in story_data:
        worksheet.write('A' + str(row), today)
        worksheet.write('B' + str(row), story['Title'])
        worksheet.write('C' + str(row), story['Link'])
        worksheet.write('D' + str(row), story['Published Date'])
        worksheet.write('E' + str(row), story['Publication'])
        worksheet.write('F' + str(row), clean_number(story['Views']))
        worksheet.write('G' + str(row), story['Earnings'])
        worksheet.write('H' + str(row), clean_time(story['Member Total Time Viewed']))
        worksheet.write('I' + str(row), clean_time(story['Average Time Viewed']))
        worksheet.write('J' + str(row), clean_number(story['Fans']))

        row += 1

    workbook.close()
    current_working_directory = os.getcwd()
    os.startfile(current_working_directory + "/" + fileName)


def write_to_gsheet(story_data):
    # Connects to your Google Sheet and writes the data into the next empty rows

    print("Connecting to your Google Sheets.")

    global today

    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    creds = Credentials.from_service_account_file(google_api_creds_filepath, scopes=scope)

    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    ss = client.open(google_sheet_workbook)
    sheet = ss.worksheet(google_sheet_workbook_sheet)

    already_uploaded_data_today = False

    gsheet_row = 1
    while sheet.cell(gsheet_row, 1).value != "":

        if sheet.cell(gsheet_row, 1).value == today:
            already_uploaded_data_today = True

        gsheet_row += 1

    if already_uploaded_data_today:
        print("************************************")
        print("Looks like you've already uploaded today's data into Google Sheets.")
        print("To upload new data, please delete today's data from your Google Sheet.")

    if already_uploaded_data_today == False:

        for story in story_data:

            try:

                row_to_update = str(('A' + str(gsheet_row) + ':J' + str(gsheet_row)))
                new_row_values = [[today, story['Title'], story['Link'], story['Published Date'], story['Publication'],
                                   clean_number(story['Views']), story['Earnings'],
                                   clean_time(story['Member Total Time Viewed']),
                                   clean_time(story['Average Time Viewed']), clean_number(story['Fans'])]]

                sheet.update(row_to_update, new_row_values)
                print("Wrote Google Sheets row for story: " + str(story['Title']))

            except:
                print("ERROR! -- Couldn't write Google Sheets row for story: " + str(story['Title']))

            gsheet_row += 1


def run_program():
    write_to_excel_boolean = True
    x = input("Would you like to write your Medium.com data to Excel?\n'Y' for Yes, 'N' for No.")
    if 'n' in x.lower():
        write_to_excel_boolean = False

    print("************************************")

    write_to_google_sheet = False

    y = input("Would you like to write your Medium.com data to Google Sheets?\n'Y' for Yes, 'N' for No.")
    if 'y' in y.lower():
        write_to_google_sheet = True

    # Login to Medium with Google
    login_google()

    # go to stats page - you are now logged in to your Medium account
    correctly_logged_in = False

    while correctly_logged_in == False:
        browser.get('https://medium.com/me/stats')
        time.sleep(time_switchpages)

        story_titles = []  # story titles
        story_links = []  # links to story data pages
        story_fans = []  # fans of each story

        # default data is on views
        soup = BeautifulSoup(browser.page_source, 'html.parser')

        divTag = soup.find_all("div", {"class": "sortableTable-title u-maxWidth450"})

        if len(divTag) == 0:
            print("ERROR! Make sure you're logged in!")
            print("When you've logged in to Medium.com, press enter.")

            enter_entry = "."

            while len(enter_entry) > 0:
                enter_entry = input("Waiting for you to press enter...")

            print("Let's try again...")

        else:
            print("************************************")
            print("Logged into Medium.com")
            print("Grabbing stats from homepage...")
            correctly_logged_in = True

    spanTag = soup.find_all("span")

    for story in divTag:
        story_titles.append(story.text)

    for div in divTag:
        story_links.append(div.find('a')['href'])

    span_id = 0
    span_text = ""
    fans = ""

    while span_id < len(spanTag):

        span_text = str(spanTag[span_id].text)

        if "fan" == span_text.lower() or "fans" == span_text.lower():
            pass

        elif "fan" in span_text.lower():
            fans = str(spanTag[span_id].text)
            story_fans.append(fans)

        span_id += 1

    story_data = []
    i = 0

    while i < len(story_titles):
        create_story_row = {'Title': story_titles[i],
                            'Published Date': "",
                            'Publication': "Self-Published",
                            'Link': story_links[i],
                            'Views': "0",
                            'Fans': story_fans[i],
                            'Earnings': "$0.00",
                            'Total Earnings': "$0.00",
                            'Average Time Viewed': "",
                            'Member Total Time Viewed': "0 min"}
        story_data.append(create_story_row)

        i = i + 1

    # Start Grabbing Data from Individual Story Pages
    for story in story_data:

        publication = "Self-Published"

        print("Getting data for story: " + story['Title'])

        browser.get(story['Link'])
        time.sleep(time_switchpages + 5)
        story_data_page_html = BeautifulSoup(browser.page_source, 'html.parser')

        story_h2Tag = story_data_page_html.find_all("h2")
        story_pTags = story_data_page_html.find_all("p")
        story_h4tags = story_data_page_html.find_all("h4")

        # Get all the Data stored in the H2 Tags
        for h2 in story_h2Tag:

            try:
                h2_text = str(h2.text)

                # Get the Story Views Data
                try:
                    story_views = int(h2_text)
                    if story_views >= 0:
                        story['Views'] = str(story_views)
                except:
                    try:
                        if h2_text.lower() != story['Title'] and "k" in h2_text.lower() and "%" not in h2_text.lower():
                            story_views = h2_text
                            story['Views'] = story_views
                    except:
                        pass

                # Get the Time Viewed Data
                try:
                    if "sec" in h2_text or "min" in h2_text or "hr" in h2_text:
                        if len(story['Average Time Viewed']) > 0:
                            story['Member Total Time Viewed'] = h2_text
                        else:
                            story['Average Time Viewed'] = h2_text
                except:
                    pass

                # Get the Earnings Data
                try:
                    if "$" in h2_text:
                        if story['Earnings'] == "$0.00":
                            story['Earnings'] = h2_text
                        else:
                            pass
                    else:
                        # It's not the Earnings data
                        pass
                except:
                    pass

            except:
                # H2 tag doesn't have text
                pass

        # Get the Published Date and Publication Data
        for h4tag in story_h4tags:

            try:
                h4_text = str(h4tag.text)

                if "published on" in h4_text.lower():
                    published_date = h4_text.lower().replace("published on ", "")

                    if "in" in published_date:
                        in_position = published_date.find("in")
                        published_date_fixed = published_date[0:in_position].lstrip().rstrip().capitalize()
                        publication = published_date[in_position + 3:].lstrip().rstrip().capitalize()

                    story['Published Date'] = published_date_fixed
                    story['Publication'] = publication
                else:
                    pass

            except:
                pass

        print("Title: " + str(story['Title']))
        print("Published Date: " + str(story['Published Date']))
        print("Publication: " + str(story['Publication']))
        print("Views: " + str(story['Views']))
        print("Fans: " + str(story['Fans']))
        print("Earnings: " + str(story['Earnings']))
        print("Member Total Time Viewed: " + str(story['Member Total Time Viewed']))
        print("Average Time Viewed: " + str(story['Average Time Viewed']))
        print("#########################")

    if write_to_excel_boolean == True:
        write_to_excel(story_data)

    if write_to_google_sheet == True:
        write_to_gsheet(story_data)

    # close the webdriver (Chrome window)

    browser.close()


print("Medium Statzler")
print("Author: Nick Canfield @ Process Zip")
print("Process Zip Website: https://processzip.com")
print("License: MIT")
print("************************************")

run_program()
print("************************************")
print("Well done! You're a Stazler!")
print("If you like this program, please be sure to clap for it and like this GitHub repo!")
print("************************************")
print("Program Completed.")
print("Exiting Program.")
