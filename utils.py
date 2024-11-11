from datetime import datetime 
import pandas as pd
import pytz
import openpyxl
import pyarrow
#-------------------------------------------------------------------

def renCol(df,a):
    '''
inputs:  df is a Series (single column dataframe)
         a is a string for new column name
         
outputs:  a new df with the new column name         
    
    '''
    df=df.rename(columns={df.columns[0]: a})
    return df


#-------------------------------------------------------------------
def symlistConv(syms):
    '''takes a list of symbols ['A', 'B', 'C'] and converts to
    a fmp multisymbol string format for their API call urls like 'A,B,C'
    '''
    syms=tuple(syms)
    return ','.join(syms)

#------------------------------------------------------------ 

import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar
from pandas.tseries.offsets import CustomBusinessDay

def ddelt(d, start=pd.Timestamp.today()):
    trading_calendar = USFederalHolidayCalendar()
    bday_us = CustomBusinessDay(calendar=trading_calendar)
    today = start
    last_business_days = pd.date_range(end=today, periods=15000, freq=bday_us)
    holidays_last_days = trading_calendar.holidays(start=last_business_days[0], end=last_business_days[-1])
    for holiday in holidays_last_days:
        last_business_days = last_business_days[last_business_days != holiday]
    return last_business_days[-d-1].strftime('%Y-%m-%d')

#--------------------------------------------------------------------------------------------------

import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar
from pandas.tseries.offsets import CustomBusinessDay


def ytd():
    trading_calendar = USFederalHolidayCalendar()
    bday_us = CustomBusinessDay(calendar=trading_calendar)
    today = pd.Timestamp.today()
    last_year_end = pd.Timestamp(year=today.year-1, month=12, day=31)
    last_year_bdays = pd.date_range(start=last_year_end, end=today, freq=bday_us)
    holidays_last_year = trading_calendar.holidays(start=last_year_end, end=today)
    for holiday in holidays_last_year:
        last_year_bdays = last_year_bdays[last_year_bdays != holiday]
    return len(last_year_bdays)


#-------------------------------------------------------------------------------------------------------

def listFrSht(sName='Sectors', fpath='helperfiles/tickerLists.xlsx'):
    '''
    sheet names:  'Sectors','ITB','RE-Retail','RE-Residential','SAAS','RE - Hotel','meme','RE - Office','RE - Ind',
                  'RE - Mort','Reg Banks','agXLI','agXLP','Ag Inputs','XHB','Trucking','Restaurants','ARKK',
                  'Leisure','Int FrtLogist','Travel','Lodg','GDX','XRT','Pot', 'position', 'SPAC', 'SHORT'
    '''
    
    return pd.read_excel(fpath, sheet_name=sName, header=None).loc[:,0].tolist()   

#----------------------------------------------------------------------------
def tvsymexp(fpath='helperfiles/Macro.txt'):
    '''enter path and file of txt file exported from trading view
       helperfiles/Macro.txt'''
   
    with open(fpath) as f:
        lines = f.readlines()
    return [i.split(':')[1] for i in lines[0].split(',')]        

#-----------------------------------------------------------------------------

import datetime as dt

from pandas.tseries.holiday import AbstractHolidayCalendar, Holiday, nearest_workday, \
    USMartinLutherKingJr, USPresidentsDay, GoodFriday, USMemorialDay, \
    USLaborDay, USThanksgivingDay


class USTradingCalendar(AbstractHolidayCalendar):
    rules = [
        Holiday('NewYearsDay', month=1, day=1, observance=nearest_workday),
        USMartinLutherKingJr,
        USPresidentsDay,
        GoodFriday,
        USMemorialDay,
        Holiday('USIndependenceDay', month=7, day=4, observance=nearest_workday),
        USLaborDay,
        USThanksgivingDay,
        Holiday('Christmas', month=12, day=25, observance=nearest_workday)
    ]


def get_trading_close_holidays(year):
    inst = USTradingCalendar()

    return inst.holidays(dt.datetime(year-1, 12, 31), dt.datetime(year, 12, 31))


def excel_columns_to_indices(columns):
    """Convert Excel-style columns (like 'A', 'D') to pandas indices (0, 3).
    to be used in module impxl
    """
    indices = [ord(col.upper()) - ord('A') for col in columns]
    return indices

import openpyxl
import pyarrow
import pandas as pd
import numpy as np

def excel_columns_to_indices(columns):
    """Convert Excel-style columns (like 'A', 'D') to pandas indices (0, 3).
    to be used in module impxl
    """
    indices = [ord(col.upper()) - ord('A') for col in columns]
    return indices

def impxl(file_path, sheet_name='Sheet1', index_col=None, excel_columns=None, column_names=None, header=None):
    """
    Extract specified columns from an Excel sheet, set the specified column as a datetime index if applicable,
    and convert other columns to numeric types if possible.

    Parameters:
    ----------
    file_path : str
        The path to the Excel file to be read.
    sheet_name : str, optional
        The name of the sheet within the Excel file to extract data from. Defaults to 'Sheet1'.
    index_col : str
        The Excel-style column reference to use as the index (e.g., 'A').
    excel_columns : list of str
        A list of Excel-style column references to extract (e.g., ['M', 'B']).
    column_names : list of str
        A list of custom column names where the first element is the index name (e.g., ['date', 'NAV', 'OtherCol']).

    header : int, list of int, None, optional
        Indicates row(s) to use as the column names. Defaults to None.

    Returns:
    -------
    pd.DataFrame
        A DataFrame with the specified index column, either as a datetime, numeric, or text index, and the other columns as numeric types if possible.
    """
    # Combine the index column with the other columns to extract
    columns_to_extract = ([index_col] if index_col else []) + (excel_columns if excel_columns else [])
    if columns_to_extract:
        column_indices = excel_columns_to_indices(columns_to_extract)
    else:
        column_indices = None
    
    # Load the specific sheet using the openpyxl engine for better compatibility with Excel files
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=column_indices if column_indices is not None else None, header=header, engine='openpyxl')
    
    try:
        # Assign custom column names
        df.columns = column_names
        df.reset_index(drop=True, inplace=True)
    except ValueError as e:
        print("ValueError:", e)
        print("\nExplanation: This error occurs when the number of columns being read from the Excel file does not match the number of column names provided. Make sure that the length of 'column_names' matches the number of columns being extracted from the Excel file.")
        raise

    # Detect the type of the index column and convert accordingly
    index_name = column_names[0]
    sample_values = df[index_name].head(10)

    # Determine if the index column should be datetime, numeric, or text
    if pd.to_numeric(sample_values, errors='coerce').notna().sum() == len(sample_values):
        # If all values can be parsed as numbers, treat the entire column as numeric
        df[index_name] = pd.to_numeric(df[index_name], errors='coerce')
        df.set_index(index_name, inplace=True)
        df.index.name = index_name
    elif pd.to_datetime(sample_values, errors='coerce').notna().sum() == len(sample_values):
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, message='Could not infer format, so each element will be parsed individually, falling back to `dateutil`')
        # If all values can be parsed as dates, treat the entire column as datetime
        df[index_name] = pd.to_datetime(df[index_name], errors='coerce')
        df.set_index(index_name, inplace=True)
        df.index.name = index_name
        # Drop rows where the index is NaT (missing dates)
        df = df[df.index.notnull()]
    else:
        # Otherwise, treat the index column as text
        df[index_name] = df[index_name].astype(str)
        df.set_index(index_name, inplace=True)
        df.index.name = index_name

    # Convert the remaining columns to numeric types if possible, but keep original values if conversion fails
    for col in column_names[1:]:
        try:
            df[col] = pd.to_numeric(df[col], errors='raise')
        except ValueError:
            df[col] = df[col].astype(str)

    return df