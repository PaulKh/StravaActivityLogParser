
from stravalib import Client
from stravalib.model import Activity
import datetime
import logging
from model import WeeksActivities, DayActivities
import json

log = logging.getLogger("APP." + __name__)

class StravaService:
    access_token :str = ""
    def __init__(self, access_token):
        self.access_token = access_token

    # This function retrieves all activities from Strava and groups them by weeks
    # It may take several calls to Strava API to retrieve all activities in chunks by 200
    def retrieve_all_activities(self, access_token) -> WeeksActivities:
        client = Client(access_token)
        all_activities = client.get_activities()
        week_activities : WeeksActivities = WeeksActivities()
        for activity in all_activities:
            if isinstance(activity, Activity):
                # logging.debug(json.dumps(activity.to_dict())) # printing all activity details
                date = str(activity.start_date)[0:10]
                date_formated = datetime.datetime.strptime(str(date), "%Y-%m-%d").date()
                    
                # group activities by weeks with monday date as a key
                monday = date_formated - datetime.timedelta(date_formated.weekday())
                if monday in week_activities:
                    date_found = False
                    for day_activity in week_activities[monday]:
                        if date_formated in day_activity:
                            day_activity[date_formated].append(activity)
                            date_found = True
                            break
                    if not date_found:
                        week_activities[monday].append(DayActivities(date_formated, [activity]))
                else:
                    week_activities.add_monday(monday,[DayActivities(date_formated, [activity])]); # [{"2020-03-03": [Activity]}]
        log.info("Success! Activities from strava retrieved, start building table")
        log.info("Total number of weeks with at least 1 training:" + str(len(week_activities.week_activities)))
        return week_activities
    
    def retrieve_full_activity(self, activity_id) -> Activity:
        client = Client(self.access_token)
        activity = client.get_activity(activity_id)
        if isinstance(activity, Activity):
            return activity
        else:
            raise Exception("Failed to retrieve activity from Strava with id: " + str(activity_id))