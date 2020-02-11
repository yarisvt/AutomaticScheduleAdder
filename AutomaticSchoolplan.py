"""
Automatically retrieve data from http://schoolplan.han.nl/SchoolplanFT_AS/ and
add it to your google calendar.
Author: Yaris van Thiel.
Version: 1.0
"""

import datetime
import time
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from credentials import USERNAME, PASSWORD
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar']


class AutomaticSchoolPlan:

    def __init__(self, _service, classname, year, week):
        """
        Creates an instance of the class.
        :param _service: The google calendar api service.
        :param classname: The class of the user.
        :param year: The year of the schoolplan the user wants.
        :param week: The week of the schoolplan the user wants.
        """
        self.all_data = []
        self.all_events = []
        self.selected_option_week = 0
        self.service = _service
        self.classname = classname
        self.year = year
        self.week = week
        self.driver = webdriver.Chrome()
        self.driver.get('http://schoolplan.han.nl/SchoolplanFT_AS/')

    def login(self):
        """
        Logs user in with a username and password, which are saved in
        credentials.py.
        """
        # wait some time to load the page
        sleep(2)

        # select the username field from xpath
        username = self.driver.find_element_by_xpath('//*[@id="username"]')
        # fill in the username
        username.send_keys(USERNAME)

        # select the passworld field from xpath
        password = self.driver.find_element_by_xpath('//*[@id="password"]')
        # fill in the password
        password.send_keys(PASSWORD)

        # press the submit button to log in
        submit = self.driver.find_element_by_xpath(
            '/html/body/div/div[2]/form/div/input[3]')
        submit.click()

    def select_schoolplan(self):
        """
        Selects the options the user has submitted.
        """
        # select the classname dropdown menu from xpath
        select = Select(self.driver.find_element_by_xpath('//*[@id="groep"]'))
        # set the classname
        select.select_by_visible_text(self.classname)

        # select the checkbox from xpath to view the hours of the lessons
        time_lesson = self.driver.find_element_by_xpath('//*[@id="lestijden"]')
        time_lesson.click()

        # select the week drowpdown menu from xpath
        select_week = Select(
            self.driver.find_element_by_xpath('//*[@id="StartWeek"]'))
        # set the week
        select_week.select_by_value(
            get_first_day_of_week(self.year, self.week, 1))
        sleep(2)

    def get_info(self):
        """
        Gets all the data from the table and appends it to a nested list. The
        list contains per hour a list with the lessons and hours.
        """
        # get the table with the data from xpath
        for tr in self.driver.find_elements_by_xpath(
                '/html/body/table/tbody/tr[2]/td/table'):
            tds = tr.find_elements_by_tag_name('td')
            data = []
            for i, td in enumerate(tds):
                # if all data from 1 hour has been added to a list. Add that
                # list to the list containing all the data. Then clear the
                # list containing data for each hour.
                if i % 7 == 0 and len(data) > 0:
                    # TODO: check for vakantie, dan niet toevoegen
                    self.all_data.append(data)
                    data = []
                data.append(td.text.strip())

    def add_events(self):
        """
        Add events, which are formatted for the google calendar api, to a list.
        """
        week_days = get_days_in_week(self.week)
        for i in range(len(self.all_data)):
            for j in range(len(self.all_data[i])):
                if self.all_data[i][j] != '':
                    try:
                        # get the lesson name
                        summary = self.all_data[i][j].split(' ')[0]
                        ind = self.all_data[i].index(self.all_data[i][j])

                        # get the start date in ISO format
                        start_date = get_iso_date(
                            week_days[ind] + self.all_data[i][
                                1])

                        # get the end date in ISO format
                        end_date = get_iso_date(
                            week_days[ind] + self.all_data[i + 1][
                                1])
                        # get the classroom in which the lesson takes place
                        classroom = self.all_data[i][j].split(' ')[2]

                        # create event
                        event = create_event(summary, start_date, end_date,
                                             classroom)
                        self.all_events.append(event)
                    except KeyError:
                        pass

    def add_to_calendar(self):
        """
        Adds all of the created events to the google calendar.
        """
        for event in self.all_events:
            # add event to google calendar
            event = self.service.events().insert(calendarId='primary',
                                                 body=event).execute()
            print('Event created: {}'.format(event.get('htmlLink')))

    def exit_page(self):
        self.driver.quit()


def authorize_calendar():
    """
    Authorizes the user's google account and saves its credentials if it's the
    first time. Otherwise it will use the already created credentials.
    :return: A built google calendar api service.
    """
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)


def create_event(summary, start_date, end_date, classroom):
    """
    Creates a event in the google calendar format.
    :param summary: The name of the lesson.
    :param start_date: The startdate of the lesson, in ISO 8601 format.
    :param end_date: The enddate of the lesson, in ISO 8601 format.
    :param classroom: The classroom in which the lesson will be taken.
    :return:
    """
    return {
        'summary': summary,
        'location': classroom,
        'start': {
            'dateTime': start_date,
            'timeZone': 'Europe/Amsterdam'
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Europe/Amsterdam'
        },
    }


def get_days_in_week(week_number):
    """
    Gets the weekdays from a given week number in format 2000-1-31 %Y-%m-%d.
    :param week_number: the week number.
    :return: a dictionary containing the weekdays of the given week number.
    """
    week_days = {}
    week = week_number - 2
    startdate = time.asctime(time.strptime('2020 %d 0' % week, '%Y %W %w'))
    startdate = datetime.datetime.strptime(startdate, '%a %b %d %H:%M:%S %Y')
    # week 0 is sunday, so start from monday (1) to friday (6)
    for i in range(1, 6):
        day = startdate + datetime.timedelta(days=i)
        week_days[i + 1] = day.strftime('%Y-%m-%d')
    return week_days


def get_first_day_of_week(year, week, whichday):
    """
    Gets a certain day of a given week and year in format 1-31-2000
    (%-m/%-d/%Y).
    :param year: The year.
    :param week: The week.
    :param whichday: The day that should be returned.
    :return: A date in format 1-31-2000.
    """
    startdate = time.asctime(time.strptime(f'{year} {week - 2} 0', '%Y %W %w'))
    startdate = datetime.datetime.strptime(startdate, '%a %b %d %H:%M:%S %Y')
    first_day = startdate + datetime.timedelta(days=whichday)
    return first_day.strftime('%-m/%-d/%Y')


def get_iso_date(date):
    """
    Gets the ISO 8601 formatted date from a regular date in format
    2020-1-3109:00.
    :param date: A date in format 2020-1-3109:00 (%Y-%m-%d%H:%M).
    :return: The ISO 8601 date format.
    """
    date = datetime.datetime.strptime(date, '%Y-%m-%d%H:%M')
    return date.isoformat()


def remove_events(service, year, week):
    """
    Removes all events during the weekdays from your google calendar
    based on a year and a week.
    :param service: The google calendar api service.
    :param year: The year of which the events should be removed.
    :param week: The week of which the events should be removed.
    """

    # get first weekday of week in ISO format
    first_day = get_first_day_of_week(year, week, 1)
    first_day = datetime.datetime.strptime(first_day,
                                           '%m/%d/%Y').isoformat() + 'Z'

    # get last weekday of week in ISO format
    last_day = get_first_day_of_week(year, week, 6)
    last_day = datetime.datetime.strptime(last_day,
                                          '%m/%d/%Y').isoformat() + 'Z'

    # get all the event between first day and last day
    events_result = service.events().list(calendarId='primary',
                                          timeMin=first_day,
                                          timeMax=last_day,
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])
    for event in events:
        # delete all the events based in event ID
        service.events().delete(calendarId='primary',
                                eventId=event['id']).execute()
        summary = event['summary']
        print(f'{summary} deleted')


# TODO: make GUI for better visualisation

if __name__ == '__main__':
    serv = authorize_calendar()
    # remove_events(serv, 2020, 10)
    schoolplan = AutomaticSchoolPlan(serv, 'BIN-2a', 2020, 8)
    schoolplan.login()
    schoolplan.select_schoolplan()
    schoolplan.get_info()
    schoolplan.add_events()
    schoolplan.add_to_calendar()
    schoolplan.exit_page()
