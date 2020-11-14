import argparse
from ExcelWriter import ExcelWriter
from ConfigsManager import ConfigsManager
from stravalib.client import Client
import datetime
from AuthorisationStrava import get_access_token

def parse_input_to_map(access_token):
    client = Client(access_token)
    all_activities = client.get_activities()
    date_activities_map = {}
    for activity in all_activities:
        date = str(activity.start_date)[0:10]
        date_formated = datetime.datetime.strptime(str(date), "%Y-%m-%d").date()
        # group activities by weeks with monday date as a key
        monday = date_formated - datetime.timedelta(date_formated.weekday())
        if monday in date_activities_map:
            date_found = False
            for day_activity in date_activities_map[monday]:
                if date_formated in day_activity:
                    day_activity[date_formated].append(activity)
                    date_found = True
                    break
            if not date_found:
                date_activities_map[monday].append({date_formated: [activity]});
        else:
            date_activities_map[monday] = [{date_formated: [activity]}]; # [{"2020-03-03": [Activity]}]
    print("Total number of weeks with at least 1 training:" + str(len(date_activities_map)))
    return date_activities_map

def main():
    #Get args
    parser = argparse.ArgumentParser(description='Params to generate result table. Table will be recreated if doesn\'t exist')
    parser.add_argument('-a', '--auth-token', dest='auth_token',
                   help='An authorisation token for strava')
    input_values = parser.parse_args()
    access_token = None
    if input_values.auth_token is not None:
        access_token = input_values.auth_token
    else:
        configs = ConfigsManager().read_configs()
        access_token = get_access_token(configs)
    
    date_activities_map = parse_input_to_map(access_token)
    ExcelWriter().create_and_fill_tables(date_activities_map)


if __name__ == '__main__':
    main()
