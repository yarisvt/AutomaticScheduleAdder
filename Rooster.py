import datetime
import time
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from credentials import EMAIL, PASSWORD
import csv
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar']


class Rooster:

    def __init__(self):
        self.all_data = []
        self.selected_option_week = 0
        self.driver = webdriver.Chrome()

    def login(self):
        self.driver.get('http://schoolplan.han.nl/SchoolplanFT_AS/')
        sleep(2)
        username = self.driver.find_element_by_xpath('//*[@id="username"]')
        username.send_keys(EMAIL)

        password = self.driver.find_element_by_xpath('//*[@id="password"]')
        password.send_keys(PASSWORD)

        submit = self.driver.find_element_by_xpath(
            '/html/body/div/div[2]/form/div/input[3]')
        submit.click()

        profile = webdriver.FirefoxProfile()
        profile.accept_untrusted_certs = True

    def get_info(self):
        select = Select(self.driver.find_element_by_xpath('//*[@id="groep"]'))

        # select by visible text
        select.select_by_visible_text('BIN-2a')

        tijd = self.driver.find_element_by_xpath('//*[@id="lestijden"]')
        tijd.click()

        select_week = Select(
            self.driver.find_element_by_xpath('//*[@id="StartWeek"]'))
        select_week.select_by_value('2/17/2020')
        sleep(2)

        select_week = Select(
            self.driver.find_element_by_xpath('//*[@id="StartWeek"]'))
        self.selected_option_week = int(select_week.first_selected_option.text.split(' ')[0])

        for tr in self.driver.find_elements_by_xpath(
                '/html/body/table/tbody/tr[2]/td/table'):
            tds = tr.find_elements_by_tag_name('td')
            data = []
            for i, td in enumerate(tds):
                if i % 7 == 0 and len(data) > 0:
                    self.all_data.append(data)
                    data = []
                data.append(td.text.strip())

    def write_file(self):
        with open('test.tsv', 'w', newline='') as file:
            writer = csv.writer(file, delimiter='\t')
            writer.writerows(self.all_data)

    def add_to_calendar(self):
        week_days = get_days_in_week(self.selected_option_week)
        for i in range(len(self.all_data)):
            for j in range(len(self.all_data[i])):
                if self.all_data[i][j] != '':
                    try:
                        summary = self.all_data[i][j].split(' ')[0]
                        ind = self.all_data[i].index(self.all_data[i][j])
                        start_date = week_days[ind] + 'T' + self.all_data[i][
                            1] + ':00+01:00'
                        end_date = week_days[ind] + 'T' + \
                                   self.all_data[i + 1][1] + ':00+01:00'
                        # print(summary, start_date, end_date)
                        add_event(summary, start_date, end_date)
                    except KeyError:
                        pass


def authorize_calendar():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)


def add_event(summary, start_date, end_date):
    # Call the Calendar API
    service = authorize_calendar()
    event = {
        'summary': summary,
        'start': {
            'dateTime': start_date,
            'timwZone': 'Netherlands/Amsterdam'
        },
        'end': {
            'dateTime': end_date,
            'timwZone': 'Netherlands/Amsterdam'
        },
    }
    # event = {
    #     'summary': 'Google I/O 2015',
    #     'start': {
    #         'dateTime': '2020-02-28T09:00:00+01:00',
    #         'timwZone': 'Netherlands/Amsterdam'
    #     },
    #     'end': {
    #         'dateTime': '2020-02-28T17:00:00+01:00',
    #         'timwZone': 'Netherlands/Amsterdam'
    #     },
    # }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))


def get_days_in_week(day_of_week):
    week_days = {}
    week = day_of_week - 2
    startdate = time.asctime(time.strptime('2020 %d 0' % week, '%Y %W %w'))
    startdate = datetime.datetime.strptime(startdate, '%a %b %d %H:%M:%S %Y')
    for i in range(1, 6):
        day = startdate + datetime.timedelta(days=i)
        week_days[i + 1] = day.strftime('%Y-%m-%d')
    return week_days


# get_days_in_week(7)

r = Rooster()
r.login()
r.get_info()
# # r.write_file()
r.add_to_calendar()
