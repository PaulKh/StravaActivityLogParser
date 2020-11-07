# StravaActivityLogParser
A simple project to generate sport activity logs based on Strava application

1. Create application on Strava. Go to settings -> My API Application -> create your application with Authorization Domain=localhost and website=https://google.com (doesn't matter) 
2. Then follow this guide https://yizeng.me/2017/01/11/get-a-strava-api-access-token-with-write-permission/
Strava requires a several step authorisation which could be done if you setup your own server with UI, but it is out of scope for this repo. Token is valid for 6 hours, so this step has to be done every time. After first time it is < 1min.

Now you should have your authorisation token and you can setup application and run it.
3. You need python setup on your machine. Tested with Python 3.8, should work with Python 2.7(but why would you use python2?)
Run:
```
python3 -m venv env
source env/bin/activate
pip install -r requirements.txt
```
This makes a full setup of all the necessary libraries for excel edit and strava api
4. Now just call 
```
main.py PUT_YOUR_AUTH_ID_HERE
```