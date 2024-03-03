
import datetime
import logging
from stravalib.model import Activity

log = logging.getLogger("APP." + __name__)
class DayActivities:
    # key is date and value is list of activities on that day
    day_activities: dict[datetime.date, list[Activity]]
    def __init__(self, date: datetime.date, activities: list[Activity]):
        self.day_activities = {date: activities}

    def __contains__(self, key: datetime.date):
        return key in self.day_activities
    
    def __getitem__(self, key) -> list[Activity]:
        return self.day_activities[key]
    
    def keys(self):
        return self.day_activities.keys()
    
    def print(self):
        for day in self.day_activities:
            logging.info(day)
            for activity in self.day_activities[day]:
                logging.info(str(activity))

class WeeksActivities:
    # key is monday date and value is list of activities per day on that week
    week_activities: dict[datetime.date, list[DayActivities]] = {}

    def __contains__(self, key: datetime.date):
        return key in self.week_activities
    
    def __getitem__(self, key) -> list[DayActivities]:
        return self.week_activities[key]
    
    def add_monday(self, date: datetime.date, day_activities: list[DayActivities]):
        self.week_activities[date] = day_activities

    def print(self):
        for week in self.week_activities:
            logging.info("Monday " +  str(week))
            for day in self.week_activities[week]:
                day.print()

    def keys(self):
        return self.week_activities.keys()
