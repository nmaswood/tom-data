import openpyxl as o
import pandas as pd
from glob import glob
from collections import namedtuple
from datetime import datetime
from dateutil.parser import parse
import pandas as pd

State = namedtuple('State', 'state start_year list_of_months')
Month = namedtuple('Month', 'year month list_of_days')
Day = namedtuple('Day', 'day percentage')

FIRST_SHEET = 'colorado_2016_2017.xlsx'

def get_all_files():

    """
    Returns a list of strings corresponding to all file names. You may have to tweak this.
    """

    return glob("*.xlsx")


def get_number_of_schools(sheet):

    """

    Returns an integer corresponding the # of schools in a file

    """

    SCHOOL_OFFSET = 6
    i = 0

    while sheet[SCHOOL_OFFSET + i][0].value is not None:
        i +=1

    return i

def create_day_dict():

    """
    Creates an empty dictionary to be filled in later
    """
    return {
        'no_school':  0,
        'first_and_last': 0,
        'half': 0,
        'other': 0,
        'school': 0
    }

def update_day_dict(d, x):

    if x is None:
        d['school'] += 1
    elif x == 'x':
        d['no_school'] += 1
    elif  x == 'h':
        d['half'] += 1
    elif x == 'f':
        d['first_and_last'] += 1
    else:
        raise Exception("What else am I missing?")


def how_many_days_off_for_date(sheet, date_offset_var, how_many_schools):

    """

    Fills in a create_day_dict dict

    """

    DATE_OFFSET_ROW = 4
    BOXES_OFFSET_COL = 6
    DATE_OFFSET_COL = 4

    PERCENTAGE_OFFSET_ROW = 46

    try:
        date = sheet[DATE_OFFSET_ROW][DATE_OFFSET_COL + date_offset_var]
    except IndexError:
        return None

    if date.value is None:
        return None

    school_day_dict = create_day_dict()

    for row_i in range(how_many_schools):
        cell = sheet[BOXES_OFFSET_COL  + row_i][DATE_OFFSET_COL + date_offset_var]
        update_day_dict(school_day_dict, cell.value)

    return school_day_dict['no_school'] / how_many_schools
    #return school_day_dict

def how_many_days_for_month(sheet, num_schools):

    days = []
    for i in range(32):

        val = how_many_days_off_for_date(sheet, i, num_schools)

        if val is None:
            break
        days.append(val)

    return days


def process_sheet(sheet):

    """
    Gets the date/ % information for a single tab
    """

    num_schools = get_number_of_schools(sheet)
    day_percentage = how_many_days_for_month(sheet,num_schools)
    month = sheet.title.split()[0]
    year = sheet.title.split()[1]


    days = [Day(idx + 1,value) for idx,value in enumerate(day_percentage)]
    month_final = Month(year, month, days)

    return month_final

def process_workbook(name):

    """
    Gets the date/ % information for the entire workbook
    """

    wb = o.load_workbook(name)

    relevant_sheets = [sheet for sheet in wb.worksheets if '_' not in sheet.title and not sheet.title.startswith("School")]

    splat = name.split("_")

    state_name = splat[0].title()

    start_year = splat[1]

    months = [process_sheet(sheet) for sheet in relevant_sheets]

    return State(state_name,start_year, months)



def str_date_format(month_tuple):

    y = month_tuple.year
    month = month_tuple.month

    s =  month + " {}, 20" + y

    return s


def process_month(month_tuple):

    base_str = str_date_format(month_tuple)
    as_dates = [parse(base_str.format(day_obj.day)) for day_obj in month_tuple.list_of_days]
    percentages = [day_obj.percentage for day_obj in month_tuple.list_of_days]

    return as_dates, percentages

def processed_workbook_to_dataframe(pwb):

    state_name = pwb.state
    list_of_months = pwb.list_of_months
    start_year =  pwb.start_year

    all_days = []
    all_percentages = []

    for month_tuple in list_of_months:
        days, percentages = process_month(month_tuple)

        all_days += days
        all_percentages += percentages

    d = {'date': all_days, state_name: all_percentages}

    df = pd.DataFrame(d)
    df = df.set_index('date')
    return df

def workbook_name_to_df(name):
    results = process_workbook(name)
    return processed_workbook_to_dataframe(results)

if __name__ == '__main__':

    res = workbook_name_to_df(FIRST_SHEET)
    print (res)