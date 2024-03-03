import argparse
import json
import logging
import time
from ExcelWriter import ExcelWriter
from ConfigsManager import ConfigsManager
from stravalib.client import Client
from stravalib.model import Activity
import datetime
from AuthorisationStrava import get_access_token
from StravaService import StravaService
from logger_setup import setup_logger
from model import WeeksActivities, DayActivities

log = logging.getLogger("APP." + __name__)


def main():
    setup_logger("logs.log", logging.DEBUG)
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
    stravaService = StravaService(access_token)
    date_activities_map = stravaService.retrieve_all_activities(access_token)
    ExcelWriter(stravaService).create_and_fill_tables(date_activities_map)
    log.info("Table updated, check activities.xlsx")


if __name__ == '__main__':
    main()
