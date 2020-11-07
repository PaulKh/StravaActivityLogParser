from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Color
import datetime

class ExcelWriter:
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
        print("First ever workout:" + str(first_workout_date))
        future_weeks = 4

        b3Value = work_sheet["B3"].value
        today = datetime.date.today()
        last_tracked_monday = today + datetime.timedelta(days=7 * future_weeks - today.weekday())
        print("Last tracked monday" + str(last_tracked_monday))
        if b3Value is not None:
            # Move existing rows down to add place for new weeks on top
            dateFromB3 = datetime.datetime.strptime(str(b3Value), "%Y-%m-%d").date()
            number_of_weeks = 0
            while dateFromB3 < last_tracked_monday:
                dateFromB3 = dateFromB3 + datetime.timedelta(days=7)
                number_of_weeks = number_of_weeks + 1
            print(str(dateFromB3) + str(type(dateFromB3)))
            if dateFromB3 != last_tracked_monday:
                raise ValueError('Error! B3 value is not Monday date')
            print(number_of_weeks)
            work_sheet.move_range("A3:Z2000", rows=number_of_weeks, cols=0)
        iteration_date = last_tracked_monday
        iterator = 3
        #fill empty dates in B column
        while first_workout_date < (iteration_date + datetime.timedelta(days=7)):
            if work_sheet["B" + str(iterator)].value is None:
                work_sheet["B" + str(iterator)] = str(iteration_date)
            iterator = iterator + 1
            iteration_date = iteration_date - datetime.timedelta(days=7)

    def __fill_columns_with_values(self, work_sheet, date_activities_map):
        iterator = 3
        while work_sheet["B" + str(iterator)].value is not None:
            value_from_cell = work_sheet["B" + str(iterator)].value
            date_from_cell = datetime.datetime.strptime(str(value_from_cell), "%Y-%m-%d").date()
            if date_from_cell in date_activities_map:
                print("week = {} workouts = {}".format(date_from_cell, len(date_activities_map[date_from_cell])))
                for activity in date_activities_map[date_from_cell]:
                    date = str(activity.start_date)[0:10]
                    date_formated = datetime.datetime.strptime(str(date), "%Y-%m-%d").date()
                    day_of_the_week = date_formated - date_from_cell
                    if work_sheet.cell(iterator, 3 + day_of_the_week.days).value is None:
                        work_sheet.cell(iterator, 3 + day_of_the_week.days).value = activity.name
            iterator = iterator + 1

    # def apply_style(self):
    #     cell.style.alignment.wrap_text=True
    def create_and_fill_tables(self, date_activities_map):
        fileName = 'activities.xlsx'
        work_book = load_workbook(fileName)
        work_sheet = self.__get_worksheet(work_book)  
        self.__fill_dates_column(work_sheet, self.__get_first_activity_date(date_activities_map))
        self.__fill_columns_with_values(work_sheet, date_activities_map)
        work_book.save(fileName)

    