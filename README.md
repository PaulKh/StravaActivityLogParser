# StravaActivityLogParser
A simple project to generate sport activity logs based on Strava application
## Features
* Scripts creates a work sheet in excel file with dates and strava activities created in this dates. 
* Every week is a new line
* First column is monday date. Next 7 columns represent days of the week
* The column after Sunday contains estimation of calories burnt, could be manually updated if needed adjustment. Note, only Ride/swim/run supported, easy to add more
* Next column contains a progress comparing to previous weeks
* Weeks are added automatically.
* Value of cell is a list of names of activities and type
* Comment of the activity contains type/speed/distance/time/calories/heart rate
* Values in cells/comments are **NOT** overriden by script. If you did any modification on values it will stay there after each new script run
## Setup guide
1. Create application on Strava. Go to settings -> My API Application -> create your application with Authorization Domain=localhost and website=https://google.com (doesn't matter). Other fields can be random
2. Then go to this page http://www.strava.com/oauth/authorize?client_id=[REPLACE_WITH_YOUR_CLIENT_ID]&response_type=code&redirect_uri=http://localhost/exchange_token&approval_prompt=force&scope=read_all,profile:read_all,activity:read_all (Put your Client id in a given place)
Normally you will be redirected to the page where you have to accept permission. This page will redirect to 
your localhost but you'll have CODE in URL that you will find in your browser address bar.
More details could be found here https://yizeng.me/2017/01/11/get-a-strava-api-access-token-with-write-permission/
3. Fill document strava_logger.cfg that you find in this repository with code that you got on the previous step, client_id and client_secret from your Strava APlication page and other fields age/sex/weight/height. That is it for configuration
4. You need python setup on your machine. Tested with Python 3.8 on unix, should work with Python 2.7(but why would you use python2?)
Run:
```
python3 -m venv env
source env/bin/activate
pip install -r requirements.txt
```
This makes a full setup of all the necessary libraries for excel edit and strava api
4. Now just call script: 
```
stravaLogGatherer.py
```
and call it every time you want to update your log with the latest data from 
