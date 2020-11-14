from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Alignment
from openpyxl.comments import Comment
from ConfigsManager import ConfigsManager
import datetime
from datetime import timedelta
from openpyxl.utils import units
import sys

class ExcelWriter:
    HEIGHT_KEY = "height"
    WEIGHT_KEY = "weight"
    AGE_KEY = "age"
    SEX_KEY = "sex"
    CALORIES_COLUMN = 10

    def __fill_sheet_with_default(self, work_sheet):
        #actually you can, but result is not quaranteed ;)
        red_font = Font(color="FF0000")
        work_sheet.cell(2, 2).value = "Don\'t modify this column/row"
        work_sheet.cell(2, 2).font = red_font
        work_sheet.cell(2, 3).value = "Monday"
        work_sheet.cell(2, 4).value = "Tuesday"
        work_sheet.cell(2, 5).value = "Wednesday"
        work_sheet.cell(2, 6).value = "Thursday"
        work_sheet.cell(2, 7).value = "Friday"
        work_sheet.cell(2, 8).value = "Saturday"
        work_sheet.cell(2, 9).value = "Sunday"
        work_sheet.cell(2, 10).value = "Calories"
        work_sheet.column_dimensions['C'].width = 25
        work_sheet.column_dimensions['D'].width = 25
        work_sheet.column_dimensions['E'].width = 25
        work_sheet.column_dimensions['F'].width = 25
        work_sheet.column_dimensions['G'].width = 25
        work_sheet.column_dimensions['H'].width = 25
        work_sheet.column_dimensions['I'].width = 25

    def __apply_style(self, work_sheet):
        iterator = 3
        while work_sheet.cell(iterator, 2).value is not None:
            for column_iterator in range(3, 12):
                work_sheet.cell(iterator, column_iterator).alignment = Alignment(wrap_text=True)
            iterator = iterator + 1

    def __get_worksheet(self, work_book):
        activityLogSheetName = "Activity Log"
        work_sheet = None 
        if activityLogSheetName not in work_book.get_sheet_names():
            print("Sheet " + activityLogSheetName + " will be created")
            work_sheet = work_book.create_sheet(activityLogSheetName)
            self.__fill_sheet_with_default(work_sheet)
        else:
            work_sheet = work_book.get_sheet_by_name(activityLogSheetName)
        return work_sheet

    
    def __get_first_activity_date(self, date_activities_map):
        first_workout_date = None
        for date in date_activities_map.keys():
            if not first_workout_date:
                first_workout_date = date
            elif date < first_workout_date:
                first_workout_date = date
        return first_workout_date

    def __fill_dates_column(self, work_sheet, first_workout_date):
        future_weeks = 4

        b3Value = work_sheet["B3"].value
        today = datetime.date.today()
        last_tracked_monday = today + datetime.timedelta(days=7 * future_weeks - today.weekday())
        if b3Value is not None:
            # Move existing rows down to add place for new weeks on top
            dateFromB3 = datetime.datetime.strptime(str(b3Value), "%Y-%m-%d").date()
            number_of_weeks = 0
            while dateFromB3 < last_tracked_monday:
                dateFromB3 = dateFromB3 + datetime.timedelta(days=7)
                number_of_weeks = number_of_weeks + 1
            if dateFromB3 != last_tracked_monday:
                raise ValueError('Error! B3 value is not Monday date')
            work_sheet.move_range("A3:Z2000", rows=number_of_weeks, cols=0)
        iteration_date = last_tracked_monday
        iterator = 3
        #fill empty dates in B column
        while first_workout_date < (iteration_date + datetime.timedelta(days=7)):
            if work_sheet["B" + str(iterator)].value is None:
                work_sheet["B" + str(iterator)] = str(iteration_date)
            iterator = iterator + 1
            iteration_date = iteration_date - datetime.timedelta(days=7)

    #Basal metabolic rate (BMR) is often used interchangeably with resting metabolic rate (RMR)
    def __compute_bmr(self):
        if self.configs[self.SEX_KEY] == "M":
            self.bmr = 88.362 + (13.397 * self.configs[self.WEIGHT_KEY]) + (4.799 * self.configs[self.HEIGHT_KEY]) - (5.677 * self.configs[self.AGE_KEY])
        else:   
            self.bmr = 447.593 + (9.247 * self.configs[self.WEIGHT_KEY]) + (3.098 * self.configs[self.HEIGHT_KEY]) - (4.330 * self.configs[self.AGE_KEY])

    #returns a number which is approximate effort estimation
    def __fill_day(self, cell, activities_for_day):
        cell_value = ''
        comment_value = ''

        #Calculate callories and effort
        calories_burnt = 0.0
        for activity in activities_for_day: # strava activity objects for a given date
            comment_value = comment_value + str(activity.type) + ": "
            cell_value = cell_value + str(activity.type) + ": " + str(activity.name) + "\n"
            if activity.average_speed is not None and activity.moving_time is not None:
                calories = 0
                if activity.type == "Run":  
                    calories = float(self.bmr * activity.average_speed * 3.6 * activity.moving_time.seconds) / (24 * 3600) 
                elif activity.type == "Swim":
                    calories = self.bmr * 8.5 * activity.moving_time.seconds / (24 * 3600) 
                elif activity.type == "Ride":
                    calories = float(self.bmr * activity.average_speed * 3.6 * activity.moving_time.seconds) / (24 * 3600 * 3) 
                calories_burnt += calories
                comment_value += "  Calories:" + str(round(calories, 0)) + "\n"
            if activity.average_speed is not None:
                comment_value += "  Average speed: " + str(round(float(activity.average_speed * 3.6), 2)) + "km/h\n"
            if activity.max_speed is not None and int(activity.max_speed) != 0:
                comment_value += "  Max speed: " + str(round(float(activity.max_speed * 3.6), 2)) + "km/h\n"
            if activity.moving_time is not None:
                comment_value += "  Moving time: " + str(activity.moving_time) + "\n"
            if activity.average_heartrate is not None:
                comment_value += "  Average heart rate: " + str(activity.average_heartrate) + "bpm\n"
            if activity.distance is not None:
                comment_value += "  Distance: " + str(round(float(activity.distance) / 1000, 3)) + "kms\n"
            comment_value += "\n"


        #Add comment
        if cell.comment is None:
            cell.comment = Comment(comment_value, "ME")
            cell.comment.width = units.points_to_pixels(250)
            cell.comment.height = units.points_to_pixels(comment_value.count('\n') * 20)
        if cell.value is None:
            cell.value = cell_value
        return calories_burnt

    def __fill_week(self, work_sheet, line_number, week_activities, monday_date):
        # try:
        calories_str = ''
        for day_activities in week_activities:
            for activity_key in day_activities.keys():
                # print(str(activity_key) + " " + str(len(day_activities[activity_key])))
                day_of_the_week = activity_key - monday_date
                cell = work_sheet.cell(line_number, 3 + day_of_the_week.days)
                activities_for_day = day_activities[activity_key]
                calories = 0
                if cell.comment is None or cell.value is None or work_sheet.cell(line_number, self.CALORIES_COLUMN) is None:
                    # If everything filled then there is nothing to do
                    calories = self.__fill_day(cell, activities_for_day)
                if calories_str == '':
                    calories_str = str(int(calories))
                else:
                    calories_str += "+" + str(int(calories)) 
                #Looks like there is a bug in recent excel that resets size of comments
                cell.comment.width = units.points_to_pixels(250)
                cell.comment.height = units.points_to_pixels(cell.comment.content.count('\n') * 20)
                
        if work_sheet.cell(line_number, self.CALORIES_COLUMN).value is None:
            work_sheet.cell(line_number, self.CALORIES_COLUMN).value = "=SUM(" + calories_str  + ")"

    def __fill_columns_with_values(self, work_sheet, date_activities_map):
        iterator = 3
        while work_sheet["B" + str(iterator)].value is not None:
            value_from_cell = work_sheet["B" + str(iterator)].value
            monday_date = datetime.datetime.strptime(str(value_from_cell), "%Y-%m-%d").date()
            if monday_date in date_activities_map:
                # print("week = {} workouts = {}".format(monday_date, len(date_activities_map[monday_date])))
                self.__fill_week(work_sheet, iterator, date_activities_map[monday_date], monday_date)
            iterator = iterator + 1

    def create_and_fill_tables(self, date_activities_map):
        self.configs = ConfigsManager().read_configs()
        self.__compute_bmr()
        fileName = 'activities.xlsx'
        work_book = load_workbook(fileName)
        work_sheet = self.__get_worksheet(work_book)  
        self.__fill_dates_column(work_sheet, self.__get_first_activity_date(date_activities_map))
        self.__fill_columns_with_values(work_sheet, date_activities_map)
        self.__apply_style(work_sheet)

        work_book.save(fileName)

    