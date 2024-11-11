#  version 1.0.0 11/11/2024  10:10 AM


from datetime import datetime 
import pandas as pd
import pytz

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


def impxl(file_path, 
          index_column,
          other_columns,
          sheet_name='Sheet1')
          header=0:
    """
    Import specified columns from an Excel sheet into a pandas DataFrame.

    Parameters
    ----------
    file_path : str
        The path to the Excel file to be read.
    index_column : str
        The name of the column to be used as the index.
    other_columns : list of str
        A list of column names to be imported along with the index column.
    sheet_name : str, optional
        The name of the sheet to read data from. Defaults to 'Sheet1'.

    Returns
    -------
    pd.DataFrame
        A DataFrame containing the specified columns, with the index column set as the DataFrame's index.
    
    Example
    -------
    >>> file_path = "C:/path/to/your/file.xlsx"
    >>> index_column = 'nameOfIndexCol'
    >>> other_columns = ['nameOfOtherColumn1', 'nameOfOtherColumn2']
    >>> sheet_name = 'nameOfTheSheet'
    >>> df = impxl(file_path, index_column, other_columns, sheet_name)
    >>> print(df.head())
    """
    # Combine the index column with the other columns to read from the Excel file
    columns_to_read = [index_column] + other_columns
    
    # Read the Excel file, specifying the header row and columns to read
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_read, header=header)
    
    # Set the index column
    df.set_index(index_column, inplace=True)
    
    # Return the DataFrame
    return df



