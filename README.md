# Automatically save schedule to google calendar
This python app will automatically retrieve your school schedule and add it to your google calendar.

## Installation steps:

### Install python package selenium
* `pip install -U selenium`
* See [selenium documentation](https://seleniumhq.github.io/selenium/docs/api/py/) for more information to setup selenium for python.

### Download chromedriver
* [Download](https://chromedriver.storage.googleapis.com/80.0.3987.16/chromedriver_linux64.zip) the linux chromedriver.
* Unzip the folder
* Move chromedriver to `/usr/local/bin`
  * `sudo mv ~/Downloads/chromedriver_linux64/chromedriver /usr/local/bin`

### Create credentials.py file with variables
```python
USERNAME = 'your_username'
PASSWORD = 'your_password'
```

### Setup google calendar API ability to your account
* Follow [step 1](https://developers.google.com/calendar/quickstart/python) to enable the Google Calendar API, and save the `credentials.json` file to your working directory.
* Install google client library `pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
`

### Run the script!
* Change class name, year and week to your likings!
```python
schoolplan = AutomaticSchoolPlan(serv, 'your_class', year, week)
```
