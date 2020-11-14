import requests
from ConfigsManager import ConfigsManager
import time

CLIENT_ID_KEY = "client_id"
CLIENT_SECRET_KEY = "client_secret"
CODE_KEY = "code"
ACCESS_TOKEN_KEY = "access_token"
REFRESH_TOKEN_KEY = "refresh_token"
EXPIRES_AT_KEY = "expires_at"
GRANT_TYPE_KEY = "grant_type"


def make_first_authorisation(configs):
    data = {}
    data[CLIENT_ID_KEY] = configs[CLIENT_ID_KEY]
    data[CLIENT_SECRET_KEY] = configs[CLIENT_SECRET_KEY]
    data[CODE_KEY] = configs[CODE_KEY]
    response = requests.post('https://www.strava.com/api/v3/oauth/token', data=data)
    authorisation_response = response.json()
    print(authorisation_response)
    configs[ACCESS_TOKEN_KEY] = authorisation_response[ACCESS_TOKEN_KEY]
    configs[REFRESH_TOKEN_KEY] = authorisation_response[REFRESH_TOKEN_KEY]
    configs[EXPIRES_AT_KEY] = authorisation_response[EXPIRES_AT_KEY]
    ConfigsManager().write_configs(configs)

def update_access_token(configs):
    print("update access token")
    data = {}
    data[CLIENT_ID_KEY] = configs[CLIENT_ID_KEY]
    data[CLIENT_SECRET_KEY] = configs[CLIENT_SECRET_KEY]
    data[GRANT_TYPE_KEY] = "refresh_token"
    data[REFRESH_TOKEN_KEY] = configs[REFRESH_TOKEN_KEY]
    response = requests.post('https://www.strava.com/api/v3/oauth/token', data=data)
    authorisation_response = response.json()
    print(authorisation_response)
    configs[ACCESS_TOKEN_KEY] = authorisation_response[ACCESS_TOKEN_KEY]
    configs[EXPIRES_AT_KEY] = authorisation_response[EXPIRES_AT_KEY]
    ConfigsManager().write_configs(configs)

def get_access_token(configs):
    if ACCESS_TOKEN_KEY not in configs or EXPIRES_AT_KEY not in configs or REFRESH_TOKEN_KEY not in configs:
        print("Login for the first time, .cfg file should contain code/client_id,client_secret")
        make_first_authorisation(configs)
    else:
        #Check if expiration date is still valid
        time_in_seconds_before_expiration = configs[EXPIRES_AT_KEY] - time.time()
        #print("token will expire in " + str(time_in_seconds_before_expiration) + " seconds")
        if time_in_seconds_before_expiration < 300:
            update_access_token(configs)
            print("Making a call to Strava to update auth token")
        else:
            print("Auth token in configs should be valid, requesting activities")
    return configs[ACCESS_TOKEN_KEY]
    
    