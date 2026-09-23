##edited 2/17/25  modified cagr and added cagrTS, added dateSlice
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression
from datetime import datetime 
from io import StringIO
from pandas.tseries.holiday import (
    AbstractHolidayCalendar,
    Holiday,
    nearest_workday,
    sunday_to_monday,
    GoodFriday,
    USMartinLutherKingJr,
    USPresidentsDay,
    USMemorialDay,
    USLaborDay,
    USThanksgivingDay,
    USFederalHolidayCalendar
)
from pandas.tseries.offsets import CustomBusinessDay
import csv
import datetime as dt
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pytz
from urllib.request import urlopen
import json
import certifi
import ssl
ssl_context = ssl.create_default_context(cafile=certifi.where())

from tqdm import notebook    #ex: for i in notebook.tqdm(range(1,100000000)):

import os
# Get the API key from the environment variable
apikey = os.getenv('FMP_API_KEY')

# Check if the API key is set
if not apikey:
    raise ValueError("API key not found. Please set the environment variable 'FMP_API_KEY'.")

##an assortment of utilities 1/30/25    http://localhost:8888/files/utils2.py?_xsrf=2%7C99db802f%7C89f4b4ca8ca6b51ddb8ecd532175eec7%7C1735847160
#-------------------------------------------------------------------import certifi
import ssl
ssl_context = ssl.create_default_context(cafile=certifi.where())
def _isin(isin):
    url = f"https://financialmodelingprep.com/api/v4/search/isin?isin={isin}&apikey="+apikey
    response = urlopen(url, context=ssl_context)
    data = response.read().decode("utf-8")
    return json.loads(data)

#-------------------------------------------------------
def renCol(df, a):
    """
inputs:  df is a Series (single column dataframe)
         a is a string for new column name
         
outputs:  a new df with the new column name         
    
    """
    df = df.rename(columns={df.columns[0]: a})
    return df

def symlistConv(syms):
    """takes a list of symbols ['A', 'B', 'C'] and converts to
    a fmp multisymbol string format for their API call urls like 'A,B,C'
    """
    syms = tuple(syms)
    return ','.join(syms)

def ddelt(d, start=None):
    """
    input: d = number of NYSE trading days as an int
           start= reference date as str 'YYYY-mm-dd' or Timestamp, default is today
           returns the date of the d-th trading day counting back from start (start itself
           is day 1 if it's a trading day) as str 'YYYY-mm-dd'
    """
    start = pd.Timestamp.today() if start is None else pd.to_datetime(start)
    start = start.normalize()
    window_start = start - pd.Timedelta(days=int(d * 1.6) + 15)
    holidays = NYSEHolidayCalendar().holidays(start=window_start, end=start)
    holidays = holidays.union(NYSE_SPECIAL_CLOSURES)
    days = pd.date_range(start=window_start, end=start, freq=CustomBusinessDay(holidays=holidays))
    return days[-d].strftime('%Y-%m-%d')

class NYSEHolidayCalendar(AbstractHolidayCalendar):
    """NYSE full-day closures.  Unlike USFederalHolidayCalendar: adds Good Friday and
    Juneteenth (2022+), drops Columbus/Veterans Day, and a Saturday New Year's Day is
    not observed on the prior Friday."""
    rules = [
        Holiday('NewYearsDay', month=1, day=1, observance=sunday_to_monday),
        USMartinLutherKingJr,
        USPresidentsDay,
        GoodFriday,
        USMemorialDay,
        Holiday('Juneteenth', month=6, day=19, start_date='2022-01-01', observance=nearest_workday),
        Holiday('IndependenceDay', month=7, day=4, observance=nearest_workday),
        USLaborDay,
        USThanksgivingDay,
        Holiday('Christmas', month=12, day=25, observance=nearest_workday),
    ]

# one-off NYSE closures (national days of mourning, Hurricane Sandy)
NYSE_SPECIAL_CLOSURES = pd.to_datetime(['2004-06-11', '2007-01-02', '2012-10-29', '2012-10-30',
                                        '2018-12-05', '2025-01-09'])

def ytd(today=None):
    """
    input: date as a string 'YYYY-mm-dd' or Timestamp, default is today
    returns:  number of NYSE trading days in that year through `today` as int
              (so px.iloc[-ytd()-1] is the prior year's last close for a US series)
    """
    today = pd.Timestamp.today() if today is None else pd.to_datetime(today)
    today = today.normalize()
    year_start = pd.Timestamp(year=today.year, month=1, day=1)
    holidays = NYSEHolidayCalendar().holidays(start=year_start, end=today)
    holidays = holidays.union(NYSE_SPECIAL_CLOSURES)
    return len(pd.date_range(start=year_start, end=today, freq=CustomBusinessDay(holidays=holidays)))

def listFrSht(sName='Sectors', fpath='helperfiles/tickerLists.xlsx'):
    """
    sheet names:  'Sectors','ITB','RE-Retail','RE-Residential','SAAS','RE - Hotel','meme','RE - Office','RE - Ind',
                  'RE - Mort','Reg Banks','agXLI','agXLP','Ag Inputs','XHB','Trucking','Restaurants','ARKK',
                  'Leisure','Int FrtLogist','Travel','Lodg','GDX','XRT','Pot', 'position', 'SPAC', 'SHORT'
    """
    return pd.read_excel(fpath, sheet_name=sName, header=None).loc[:, 0].tolist()

def tvsymexp(fpath='helperfiles/Macro.txt'):
    """enter path and file of txt file exported from trading view
       helperfiles/Macro.txt"""
    with open(fpath) as f:
        lines = f.readlines()
    return [i.split(':')[1] for i in lines[0].split(',')]

def get_trading_close_holidays(year=datetime.today().year):
    """
    inputs:  year as int. 
    returns: a datetime index of holiday dates
    """
    inst = USTradingCalendar()
    return inst.holidays(dt.datetime(year - 1, 12, 31), dt.datetime(year, 12, 31))

def impxl(file_path, index_column, other_columns, sheet_name='Sheet1', header=0):
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
    columns_to_read = [index_column] + other_columns
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_read, header=header)
    df.set_index(index_column, inplace=True)
    return df

def make_clickable(val):
    return f'<a target="_blank" href="{val}">{val}</a>'

def lr(X, Y, _print=True):
    """
General Linear Regression function

inputs
    X: series or dataframe columns for the independent variable
    Y: series or dataframe columns for the dependent variable
    _print: (Bool) True returns a printout of the regression coefficients and the coefficients
                   False only returns the coefficients: slope, intercept, r_value, p_value, std_err
                   
outputs: slope, intercept, r_value, p_value, std_err                   
    
    """
    from scipy.stats import linregress
    slope, intercept, r_value, p_value, std_err = linregress(X, Y)
    if _print:
        print(f'eq:  y = {slope: .3f}x + {intercept:.3f}')
        print(f'std er:  {std_err: .3f}')
        print(f'r value:  {r_value:.3f}')
        print(f'p value: {p_value: .3f}')
    return (slope, intercept, r_value, p_value, std_err)

def lrPoly(series: pd.Series, 
                     degrees: int = 1, 
                     return_coeffs: bool = False):
    """
    Create a DataFrame with original series and polynomial fit predictions,
    optionally returning model coefficients and last predicted value.

    Parameters:
    -----------
    series : pd.Series
        Input time series with datetime index and named values
    degrees : int, optional (default=1)
        Degree of polynomial fit (1=linear, 2=quadratic, 3=cubic)
    return_coeffs : bool, optional (default=False)
        If True, returns a tuple with DataFrame and dictionary containing
        model coefficients and last predicted value

    Returns:
    --------
    pd.DataFrame or Tuple[pd.DataFrame, dict]
        - If return_coeffs=False: DataFrame with original values and predictions
        - If return_coeffs=True: Tuple containing:
            - DataFrame with original values and predictions
            - Dictionary with:
                - 'coefficients': array of polynomial coefficients (highest degree first)
                - 'coeff_explanation': str explaining the coefficients
                - 'last_value': last predicted value

    Raises:
    -------
    ValueError
        If degrees is not 1, 2, or 3
    TypeError
        If series is not a pandas Series

    Examples:
    --------
    >>> dates = pd.date_range('2023-01-01', periods=5)
    >>> s = pd.Series([1, 2, 3, 4, 5], index=dates, name='values')
    >>> df = polynomial_fit_df(s, degrees=1)
    >>> df_coeffs, info = polynomial_fit_df(s, degrees=2, return_coeffs=True)
    """
    # Input validation
    if not isinstance(series, pd.Series):
        raise TypeError("Input must be a pandas Series")
    if degrees not in [1, 2, 3]:
        raise ValueError("degrees must be 1, 2, or 3")
    
    # Create x values as 1 to n sequence
    x = np.arange(1, len(series) + 1).reshape(-1, 1)
    
    # Create polynomial features
    poly = PolynomialFeatures(degree=degrees)
    x_poly = poly.fit_transform(x)
    
    # Create and fit polynomial regression model
    model = LinearRegression()
    model.fit(x_poly, series.values)
    
    # Calculate predicted y values
    y_pred = model.predict(x_poly)
    
    # Determine column name based on degree
    pred_col_name = {
        1: 'Linear_Prediction',
        2: 'Quadratic_Prediction',
        3: 'Cubic_Prediction'
    }[degrees]
    
    # Create DataFrame with original values and predictions
    df = pd.DataFrame({
        series.name: series,
        pred_col_name: y_pred
    }, index=series.index)
    
    if return_coeffs:
        # Get coefficients (excluding the bias term from poly, using intercept directly)
        # Reverse to match polynomial convention: highest degree first
        coeffs = model.coef_[1:][::-1]  # Skip the bias term (first element)
        coeffs = np.append(coeffs, model.intercept_)  # Add intercept as last term
        
        # Explanation text based on degree
        coeff_explanation = {
            1: "Coefficients represent: [x term, intercept] for equation y = ax + b",
            2: "Coefficients represent: [x² term, x term, intercept] for equation y = ax² + bx + c",
            3: "Coefficients represent: [x³ term, x² term, x term, intercept] for equation y = ax³ + bx² + cx + d"
        }[degrees]
        
        return_info = {
            'coefficients': coeffs,
            'coeff_explanation': coeff_explanation,
            'last_value': y_pred[-1]
        }
        return df, return_info
    
    return df

def remove_chars(lst):
    """
cleans up FMP symbols from screen.  remove symbols with "." (foreigh stocks)
and symboils with "-" (preferred stocks)
    """
    return [item for item in lst if '-' not in item and '.' not in item]

from datetime import datetime

def cagr(start_date, end_date, start_value, end_value):
    """
    Calculate the Compound Annual Growth Rate (CAGR).

    CAGR is the rate at which an investment grows annually over a specified period, 
    assuming compounding occurs at a constant rate.

    Parameters:
        start_date (str): The start date in the format 'YYYY-MM-DD'.
        end_date (str): The end date in the format 'YYYY-MM-DD'.
        start_value (float): The initial value of the investment or metric.
        end_value (float): The final value of the investment or metric.

    Returns:
        float: The CAGR as a decimal (e.g., 0.10 for 10%).

    Raises:
        ValueError: If `start_value` is zero (to prevent division errors).
        ValueError: If the number of years between `start_date` and `end_date` is zero.

    Example:
        >>> cagr('2015-01-01', '2020-01-01', 1000, 2000)
        0.1487  # (14.87% annual growth)
    """
    start_date = datetime.strptime(start_date, '%Y-%m-%d')
    end_date = datetime.strptime(end_date, '%Y-%m-%d')
    num_years = (end_date - start_date).days / 365.25  # Account for leap years

    if start_value == 0 or num_years == 0:
        raise ValueError('Start value and number of years must be non-zero')

    cagr = (end_value / start_value) ** (1 / num_years) - 1
    return cagr


def cagrTS(data):
    """
    Calculate CAGR (Compound Annual Growth Rate) for a given Pandas Series or all columns in a DataFrame.
    
    Parameters:
        data (pd.Series or pd.DataFrame): A time-series Pandas Series or DataFrame with a datetime index.

    Returns:
        float (if Series) or pd.Series (if DataFrame) with CAGR values.
    """
    data = data.dropna()  # Ensure no missing values
    if data.empty or len(data) < 2:
        return None  # Not enough data points to calculate CAGR
    
    # Calculate number of years
    n_years = (data.index[-1] - data.index[0]).days / 365.25  # Account for leap years
    
    if isinstance(data, pd.Series):
        # If input is a Series, calculate CAGR for that Series
        V_i, V_f = data.iloc[0], data.iloc[-1]
        return (V_f / V_i) ** (1 / n_years) - 1 if V_i > 0 else None
    
    elif isinstance(data, pd.DataFrame):
        # If input is a DataFrame, calculate CAGR for each column
        cagr_dict = {}
        for col in data.columns:
            col_data = data[col].dropna()
            if len(col_data) < 2:
                cagr_dict[col] = None
            else:
                V_i, V_f = col_data.iloc[0], col_data.iloc[-1]
                cagr_dict[col] = (V_f / V_i) ** (1 / n_years) - 1 if V_i > 0 else None
        return pd.Series(cagr_dict)

def string_to_csv(input_string, csv_file_path):
    """
    ****Only works on text****
    1 - copy a column of text from a spreadsheet and paste into a notebook
    2 - surround text with triple quotes and name...  'text=pasted text
    3 - input: 'text' output: 'text.csv'
    
    """
    fake_file = StringIO(input_string)
    with open(csv_file_path, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerows([[line.strip()] for line in fake_file])

def parse_multi_line_string(input_string):
    """
    Parse a multi-line string into a list of individual lines.

    This function takes a string containing multiple lines, removes leading and trailing
    whitespace, and splits it into a list where each element is a line from the input.

    Parameters:
        input_string (str): The multi-line string to parse.

    Returns:
        list: A list of strings, where each string is a line from the input with leading
              and trailing whitespace preserved within each line.

    Notes:
        - Empty lines in the input are preserved as empty strings in the output list.
        - Leading and trailing whitespace of the entire string is removed before splitting.
        - Uses the newline character ('\n') as the delimiter.

    Examples:
        >>> parse_multi_line_string('line1\nline2\nline3')
        ['line1', 'line2', 'line3']
        >>> parse_multi_line_string('  line1  \n\n  line3  ')
        ['  line1  ', '', '  line3  ']
        >>> parse_multi_line_string('single line')
        ['single line']
    """
    return input_string.strip().split('\n')
    
def csv_to_list(csv_file_path):
    """
    input: csv file path to a one column csv
    ouput: a list
    """
    result_list = []
    with open(csv_file_path, 'r') as csvfile:
        csv_reader = csv.reader(csvfile)
        for row in csv_reader:
            result_list.append(row[0])
    return result_list

def tsConvert(timestamp_ms):
    """converts a ts object like 1677790926832 into
    a datetime object"""
    timestamp_s = timestamp_ms / 1000
    datetime_obj = datetime.fromtimestamp(timestamp_s)
    return datetime_obj

def splitSymWeights(varForData):
    """For importing citrindex from excel into a notebook 
       and converting 
       input:  a 2 column paste from excel like
           WMS	0.17%
           ACM	0.06%
           ALFA	0.18%
           ATMU	0.15%
           BMI	0.33%
       returns:  a tuple of 2 lists... string:symbols and float:weights 
    
        """
    lines = data.strip().split('\n')
    symbols = []
    weights = []
    for line in lines:
        symbol_part, weight_part = line.split()
        symbols.append(symbol_part)
        weights.append(float(weight_part.strip('%')))
    return (symbols, weights)

def histVol(ser, lbk=60):
    """
    input: a series of  prices
    lbk: the number of days to calculate
    returns:  annualized hist vol from a price series using:
    (np.std(np.log(ser/ser.shift()), axis=0)*252**.5
    which is sdt of log returns multiplied by the sqrt of 252
    
    """
    ser = ser[-lbk:]
    return np.round(np.std(np.log(ser / ser.shift()), axis=0) * 252 ** 0.5 * 100, 1)

def beta(df, mkt='SPY', lbk=100):
    """
    inputs:
    df: a single column dataframe of prices
    mkt: str symbol
    lbk: int days
    returns: float"""
    dff = fmp_price(mkt, start=utils.ddelt(len(df) + 1))
    newdf = pd.concat([df, dff], axis=1).dropna()
    x = newdf.iloc[:, 1].tolist()[-(lbk + 1):]
    y = newdf.iloc[:, 0].tolist()[-(lbk + 1):]
    df = pd.DataFrame(list(zip(x, y)), columns=['x', 'y'])
    df = np.log(df / df.shift()).dropna()
    cov = df.cov()
    var = df['x'].var()
    m = cov / var
    print('lookback = ', len(df))
    return np.round(m.iloc[0, 1], 2)

def y_fmt(y, pos):
    if y == 0:
        return '0'
    exp = int(np.log10(abs(y)))
    if exp >= 9:
        return '{:.1f}B'.format(y / 1000000000.0)
    elif exp >= 6:
        return '{:.1f}M'.format(y / 1000000.0)
    elif exp >= 3:
        return '{:.1f}K'.format(y / 1000.0)
    else:
        return '{:.0f}'.format(y)

def myPlot(data, kind='line'):
    fig, ax = plt.subplots(figsize=(10, 6))
    data.plot(kind=kind, ax=ax)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(y_fmt))
    plt.legend()
    plt.grid(alpha=0.5, linestyle='--')



def dateSlice(df, years=5):
    
    """
    Slice a datetime-indexed DataFrame to keep only the most recent `years` of data.

    Parameters:
        df (pd.DataFrame): A Pandas DataFrame with a DatetimeIndex.
        years (int, optional): Number of years to keep (default is 5).

    Returns:
        pd.DataFrame: A DataFrame containing only the most recent `years` of data.
    """
    if df.empty or not isinstance(df.index, pd.DatetimeIndex):
        raise ValueError("The DataFrame must have a DatetimeIndex and cannot be empty.")

    # Determine the most recent date in the DataFrame
    latest_date = df.index.max()
    
    # Calculate the cutoff date
    cutoff_date = latest_date - pd.DateOffset(years=years)

    # Slice the DataFrame to keep only data from the last `years`
    return df[df.index >= cutoff_date]

#---------------------------------------------------------------------------------------------------
def get_symbols_from_isins1(isin_list):
    """
    Process a list of ISINs and return a DataFrame with ISIN as index and symbol/name columns.
    
    Args:
        isin_list (list): List of ISIN strings
        
    Returns:
        pandas.DataFrame: DataFrame with ISIN as index and symbol/name columns
    """
    # Initialize lists to store data
    symbols = []
    names = []
    
    for isin in notebook.tqdm(isin_list):
        try:
            result = _isin(isin)
            # Check if the result is a list and has at least one item
            if isinstance(result, list) and len(result) > 0:
                # Find the first dictionary where 'symbol' contains a '.'
                selected_dict = None
                for item in result:
                    if isinstance(item, dict) and 'symbol' in item and '.' in str(item['symbol']):
                        selected_dict = item
                        break
                
                if selected_dict:
                    # Extract the 'symbol' and 'name' from the selected dictionary
                    symbol = selected_dict.get('symbol')
                    name = selected_dict.get('companyName')
                    symbols.append(symbol if symbol else pd.NA)
                    names.append(name if name else pd.NA)
                else:
                    # If no dictionary with '.' in symbol is found, use the first dictionary or NA
                    first_dict = result[0]
                    symbol = first_dict.get('symbol')
                    name = first_dict.get('companyName')
                    symbols.append(symbol if symbol else pd.NA)
                    names.append(name if name else pd.NA)
                    print(f"No symbol with '.' found for ISIN: {isin}, using first result")
            else:
                symbols.append(pd.NA)
                names.append(pd.NA)
                print(f"No data found for ISIN: {isin}")
        except Exception as e:
            symbols.append(pd.NA)
            names.append(pd.NA)
            print(f"Error processing ISIN {isin}: {str(e)}")
    
    # Create DataFrame
    df = pd.DataFrame({
        'name': names,
        'isin': isin_list
    }, index=symbols)
    
    return df
#------------------------------------------------------------------------------------------------
def get_symbols_from_isins(isin_list):
    """
    Process a list of ISINs and return a DataFrame with ISIN as index and symbol/name columns.
    
    Args:
        isin_list (list): List of ISIN strings
        
    Returns:
        pandas.DataFrame: DataFrame with ISIN as index and symbol/name columns
    """
    # Initialize lists to store data
    symbols = []
    names = []
    
    for isin in notebook.tqdm(isin_list):
        try:
            result = _isin(isin)
            # Check if the result is a list and has at least one item
            if isinstance(result, list) and len(result) > 0:
                # Get the country code from the first two letters of ISIN
                isin_country = isin[:2]
                
                # Find the first dictionary where 'country' matches the ISIN country code
                selected_dict = None
                for item in result:
                    if (isinstance(item, dict) and 
                        'country' in item and 
                        item['country'] == isin_country):
                        selected_dict = item
                        break
                
                if selected_dict:
                    # Extract the 'symbol' and 'name' from the selected dictionary
                    symbol = selected_dict.get('symbol')
                    name = selected_dict.get('companyName')
                    symbols.append(symbol if symbol else pd.NA)
                    names.append(name if name else pd.NA)
                else:
                    # If no dictionary with matching country is found, use the first dictionary or NA
                    first_dict = result[0]
                    symbol = first_dict.get('symbol')
                    name = first_dict.get('companyName')
                    symbols.append(symbol if symbol else pd.NA)
                    names.append(name if name else pd.NA)
                    print(f"No symbol with matching country '{isin_country}' found for ISIN: {isin}, using first result")
            else:
                symbols.append(pd.NA)
                names.append(pd.NA)
                print(f"No data found for ISIN: {isin}")
        except Exception as e:
            symbols.append(pd.NA)
            names.append(pd.NA)
            print(f"Error processing ISIN {isin}: {str(e)}")
    
    # Create DataFrame
    df = pd.DataFrame({
        'name': names,
        'isin': isin_list
    }, index=symbols)
    
    return df
#----------------------------------------------------------------------------------------------
def parseBBsymbols(data, output_type='list', extra_columns=None):
    """
    Transform a string of space-separated symbol-country pairs and optional additional columns
    into a list of lists or list of dictionaries.

    Parameters:
        data (str): A multi-line string with space-separated symbol-country pairs and optional
            additional columns (e.g., 'BZU IM' or 'BZU IM 5.88').
        output_type (str, optional): Output format, either 'list' or 'dict'. Defaults to 'list'.
        extra_columns (list, optional): List of names for additional columns beyond symbol-country pair.
            If None, defaults to ['col2', 'col3', ...]. Length must match number of extra columns.

    Returns:
        list: If output_type='list', returns [[symbol, country, col2, col3, ...], ...].
        list: If output_type='dict', returns [{"symbol": symbol, "country": country, "col2": col2, "col3": col3, ...}, ...].

    Raises:
        ValueError: If output_type is not 'list' or 'dict', or if extra_columns length doesn't match.
    """
    # Validate output_type
    if output_type not in ['list', 'dict']:
        raise ValueError("output_type must be 'list' or 'dict'")

    # Split the string into lines and filter out empty lines
    lines = [line.strip() for line in data.split('\n') if line.strip()]
    
    # Process based on output_type
    if output_type == 'list':
        result = []
        for line in lines:
            parts = line.split()  # Split on any whitespace (spaces or tabs)
            if len(parts) < 2:  # Need at least symbol and country
                raise ValueError(f"Invalid line, must have at least symbol and country: {line}")
            # The first two parts are the symbol and country
            symbol, country = parts[0], parts[1]
            # Include the remaining parts as is
            remaining_parts = parts[2:] if len(parts) > 2 else []
            result.append([symbol, country] + remaining_parts)
        return result
    else:  # output_type == 'dict'
        result = []
        for line in lines:
            parts = line.split()  # Split on any whitespace (spaces or tabs)
            if len(parts) < 2:  # Need at least symbol and country
                raise ValueError(f"Invalid line, must have at least symbol and country: {line}")
            # The first two parts are the symbol and country
            symbol, country = parts[0], parts[1]
            # Prepare the dictionary
            stock_dict = {"symbol": symbol, "country": country}
            # Handle additional columns
            extra_cols_count = len(parts) - 2  # Number of parts after symbol and country
            if extra_columns is not None:
                if extra_cols_count != len(extra_columns):
                    raise ValueError(
                        f"Length of extra_columns ({len(extra_columns)}) does not match "
                        f"number of extra columns in line ({extra_cols_count}): {line}"
                    )
                col_names = extra_columns
            elif extra_cols_count > 0:
                col_names = [f"col{i+3}" for i in range(extra_cols_count)]
            else:
                col_names = []  # No additional columns
            
            # Add the additional columns to the dictionary
            for i, col_name in enumerate(col_names):
                stock_dict[col_name] = parts[i + 2]  # Add as is, starting after symbol and country
            
            result.append(stock_dict)
        return result
#----------------------------------------------------------------------------------------------
import csv
import os
from datetime import datetime
import re

def portFromFMP(var_name='portfolio', 
                csv_path=r"C:\Users\bblou\OneDrive\Desktop\Temp\ExpPort.csv", 
                output_path=r"C:\Users\bblou\OneDrive\Python\Notebooks\portfolios.py"):
    
   
    # Expected number of columns
    expected_columns = 9
    
    # Read the CSV and create the dictionary
    stocks = []
    try:
        with open(csv_path, 'r') as file:
            csv_reader = csv.reader(file)
            for row in csv_reader:
                if len(row) != expected_columns:
                    raise ValueError(
                        f"CSV row has {len(row)} columns, expected {expected_columns}: {row}"
                    )
                date, symbol, fmpSymbol, shares, cost, tradeAmt, portView, security, type_ = row
                if not symbol.strip():
                    print(f"Skipping row with empty symbol: {row}")
                    continue
                try:
                    date_obj = datetime.strptime(date, "%m/%d/%Y")
                    stock_dict = {
                        "date": date_obj,
                        "symbol": symbol,
                        "fmpSymbol": fmpSymbol,
                        "shares": float(shares),
                        "cost": float(cost),
                        "tradeAmt": float(tradeAmt),
                        "portView": portView,
                        "security": security,
                        "type": type_
                    }
                    stocks.append(stock_dict)
                except ValueError as e:
                    print(f"Skipping row with invalid data: {row} - Error: {str(e)}")
    except FileNotFoundError:
        print(f"Error: CSV file not found at {csv_path}")
        return
    except ValueError as e:
        print(f"Error: {str(e)}")
        return
    except Exception as e:
        print(f"Error reading CSV file: {str(e)}")
        return

    if not stocks:
        print("Warning: No valid stock data was processed. The dictionary will be empty.")
        return
    
    # Generate the new dictionary content
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dict_content = f"# Added on {current_date}\n{var_name} = [\n"
    for stock in stocks:
        stock_str = stock.copy()
        stock_str["date"] = stock["date"].strftime("%m/%d/%Y")
        dict_content += f"    {stock_str},\n"
    dict_content += "]\n"

    # Handle writing to the file
    try:
        if os.path.exists(output_path):
            # Read existing content
            with open(output_path, 'r') as file:
                existing_content = file.read()

            # Check if var_name exists using regex to find the exact dictionary definition
            pattern = rf"(# Added on .*\n)?{re.escape(var_name)}\s*=\s*\[.*?\](?=\n\n|\Z)"
            match = re.search(pattern, existing_content, re.DOTALL)
            
            if match:
                # Ask user if they want to replace the existing dictionary
                response = input(f"A dictionary named '{var_name}' already exists in {output_path}. Replace it? (yes/no): ").strip().lower()
                if response in ('yes', 'y'):
                    # Replace only the matching dictionary
                    new_content = re.sub(pattern, dict_content.rstrip(), existing_content, count=1, flags=re.DOTALL)
                    with open(output_path, 'w') as file:
                        file.write(new_content)
                    print(f"Dictionary '{var_name}' replaced in {output_path}")
                else:
                    print(f"Operation cancelled. Dictionary '{var_name}' not modified.")
                    return
            else:
                # Append the new dictionary
                with open(output_path, 'a') as file:
                    file.write(f"\n\n{dict_content}")
                print(f"Dictionary '{var_name}' appended to {output_path}")
        else:
            # Create new file
            header = "# Collection of stock dictionaries\n\n"
            with open(output_path, 'w') as file:
                file.write(header + dict_content)
            print(f"New file created with dictionary '{var_name}' at {output_path}")
    except Exception as e:
        print(f"Error writing to file: {str(e)}")

#-------------------------------------------------------------------------

def parse_multi_line_string2(input_string, split_columns=True, expected_columns=None):
    """
    Parse a tab-delimited multi-line string into a list of lines or a list of column lists.

    This function takes a multi-line string from an Excel sheet with tab-delimited columns,
    removes leading and trailing whitespace, and either returns the lines as strings or
    splits each line into columns based on tabs.

    Parameters:
        input_string (str): The multi-line string to parse (tab-delimited columns expected).
        split_columns (bool, optional): If False (default), returns a list of lines as strings.
            If True, splits each line into tab-delimited columns.
        expected_columns (int, optional): The expected number of columns per line (default None).
            If provided, validates that each line has this many parts after splitting.

    Returns:
        list: If split_columns is False, a list of strings where each string is a line.
        list: If split_columns is True, a list of lists where each inner list contains
              the tab-delimited columns of a line.

    Raises:
        ValueError: If expected_columns is provided but a line does not match that number
            of columns after splitting.
    """
    # Remove leading and trailing whitespace from the entire string
    input_string = input_string.strip()

    # Split the string into lines
    lines = input_string.split('\n')

    # If split_columns is False, return the list of lines as is
    if not split_columns:
        return lines

    # Process each line into tab-delimited columns
    result = []
    for line in lines:
        line = line.strip()  # Remove leading/trailing whitespace from each line
        parts = line.split('\t')  # Split on tabs only

        # Validate the number of columns if expected_columns is specified
        if expected_columns is not None:
            if len(parts) != expected_columns:
                raise ValueError(f"Line '{line}' has {len(parts)} columns, expected {expected_columns}")

        result.append(parts)

    return result

#-----------------------------------------------------------------------------
def add_summary_row(df, column_methods):
    """
    Adds a summary row to a DataFrame with specified aggregation methods for each column.
    Counts are returned as integers, and non-summarized columns are left blank.
    
    Parameters:
    df (pandas.DataFrame): Input DataFrame
    column_methods (dict): Dictionary with column names as keys and methods 
                          ('sum', 'mean', or 'count') as values
                          Example: {'A': 'sum', 'B': 'mean', 'C': 'count'}
    
    Returns:
    pandas.DataFrame: DataFrame with summary row appended
    """
    # Validate input is a dictionary
    if not isinstance(column_methods, dict):
        raise ValueError("column_methods must be a dictionary")
    
    # Validate columns exist in DataFrame
    invalid_cols = [col for col in column_methods.keys() if col not in df.columns]
    if invalid_cols:
        raise ValueError(f"Columns not found in DataFrame: {invalid_cols}")
    
    # Validate methods
    valid_methods = ['sum', 'mean', 'count']
    invalid_methods = [m for m in column_methods.values() if m not in valid_methods]
    if invalid_methods:
        raise ValueError(f"Methods must be one of {valid_methods}. Found: {invalid_methods}")
    
    # Create a copy of the DataFrame to avoid modifying the original
    result_df = df.copy()
    
    # Calculate summary values for each column based on specified method
    summary_values = {}
    for col, method in column_methods.items():
        if method == 'sum':
            summary_values[col] = result_df[col].sum()
        elif method == 'mean':
            summary_values[col] = result_df[col].mean()
        else:  # count
            summary_values[col] = int(result_df[col].count())  # Convert count to integer
    
    # Create summary row with empty strings for non-specified columns
    summary_row = pd.Series(index=result_df.columns, dtype=object)
    summary_row[:] = ''  # Initialize all values as empty strings
    
    # Fill in calculated summary values
    for col, value in summary_values.items():
        summary_row[col] = value
    
    # Add generic 'summary' as index
    summary_row.name = 'summary'
    
    # Append summary row to DataFrame
    result_df = pd.concat([result_df, summary_row.to_frame().T])
    
    return result_df
