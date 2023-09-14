def sqlcmd_query_to_df(server,user,password,query):
    """
    This function executes a SQL query on a remote server and returns the result as a Pandas Dataframe.
    On this approach, this is achieved running an Invoke-SQLCMD script on Windows Powershell as a subprocess.
    Therefore, when this function is called, you will see a Powershell window running on the background.
    
    Parameters:
    ----------------------
        server: (string) SQL Server Server/Instance name to establish connection.
        user: (string) login ID.
        password: (string) Password associated with the username.
        query: (string) SQL query to execute.
        
    Returns:
    ----------------------
        Pandas.Dataframe: A dataframe containing the query result.

    Example:
    ----------------------
        >>> server = '192.0.1.1\XXXXXX'
        >>> user = 'username'
        >>> password = 'pass123'
        >>> query = 'SELECT name from master.sys.databases'
        >>> df = sqlcmd_query_to_df(server,user,password,query)
    
    """
    import subprocess as sp
    import pandas as pd
    import tempfile
    
    temp = tempfile.gettempdir()+'\\temp.csv'
    pw = sp.Popen("where powershell", stdout=sp.PIPE).communicate()[0].decode(errors = 'ignore').rstrip()
    sp.call(pw+" Invoke-sqlcmd -Query "+"'"+query+"'"+" -Server "+server+" -User "+user+
            " -Password "+password+" |Export-Csv -NoTypeInformation -Path "+temp+" -Encoding UTF8")
    output = pd.read_csv(temp)
    
    return output


def append_df_to_google_sheets(df,service_account_key,sheet_key,sheet_name):
    """
    This function appends the records from the provided DataFrame to a Google Sheets document.
    It does not clear the existing data in the worksheet; instead, it appends the new data as new rows below the existing content.
    This is achieved using Google Sheets API with a service account and its json key which are previously created on a Google Cloud Platform project.

    Parameters:
    -----------
        df (pandas.DataFrame): The DataFrame containing the data to be appended.
        service_account_key (str): The path to the service account key JSON file.
        sheet_key (str): The unique key of the Google Sheets document.
        sheet_name (str): The name of the worksheet in the document.

    Returns:
    --------
        None

    Example:
    --------
        >>> import pandas as pd
        >>> df = pd.DataFrame({'Numbers': [1, 2, 3], 'Squares': ['1', '4', '9']})
        >>> service_account_key = 'path/to/service_account_key.json'
        >>> sheet_key = 'your_sheet_key'
        >>> sheet_name = 'Sheet1'
        >>> append_df_to_google_sheets(df, service_account_key, sheet_key, sheet_name)
    """
    from google.oauth2.service_account import Credentials
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    import gspread

    scopes = ['https://www.googleapis.com/auth/spreadsheets',
              'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(service_account_key, scopes=scopes)
    gc = gspread.authorize(credentials)
    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)
    gs = gc.open_by_key(sheet_key)
    worksheet = gs.worksheet(sheet_name)
    df=df.fillna('')
    rec= len(df)
    values = df.values.tolist()
    gs.values_append(sheet_name, {'valueInputOption': 'USER_ENTERED'}, {'values': values})

    return print(f'Succesfully added {rec} records into google sheet {sheet_key}, See at: https://docs.google.com/spreadsheets/d/{sheet_key}')    


def df_to_google_sheets(df,service_account_key,sheet_key,sheet_name):
    """
    Replace the contents of a Google Sheets worksheet with a DataFrame.
    This function clears the existing data in the specified worksheet of a Google Sheets document and replaces it
    with the data from the provided DataFrame. The entire content of the worksheet, including headers, will be
    replaced.

    Parameters:
    -----------
        df (pandas.DataFrame): The DataFrame containing the data.
        service_account_key (str): The path to the service account key JSON file.
        sheet_key (str): The unique key of the Google Sheets document.
        sheet_name (str): The name of the worksheet in the document.

    Returns:
    --------
        None

    Example:
    --------
        >>> import pandas as pd
        >>> df = pd.DataFrame({'Numbers': [1, 2, 3], 'Squares': ['1', '4', '9']})
        >>> service_account_key = 'path/to/service_account_key.json'
        >>> sheet_key = 'your_sheet_key'
        >>> sheet_name = 'Sheet1'
        >>> df_to_google_sheets(df, service_account_key, sheet_key, sheet_name)
    """
    from google.oauth2.service_account import Credentials
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    from gspread_dataframe import set_with_dataframe
    import gspread

    scopes = ['https://www.googleapis.com/auth/spreadsheets',
              'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(service_account_key, scopes=scopes)
    gc = gspread.authorize(credentials)
    gauth = GoogleAuth()
    drive = GoogleDrive(gauth)
    df=df.fillna('')
    rec = len(df)
    worksheet = gc.open_by_key(sheet_key).worksheet(sheet_name)
    worksheet.clear()
    set_with_dataframe(worksheet=worksheet, dataframe=df, include_index=False, include_column_header=True)
    
    return print(f'Succesfully cleared document and written {rec} records into google sheet {sheet_key}, See at: https://docs.google.com/spreadsheets/d/{sheet_key}') 

def query_to_df(server,user,password,query,database=None,driver='SQL Server'):
    """
    This function executes a SQL query and return the result as a Pandas DataFrame.
    On this approach, the database connection is established using ODBC protocol with pyodbc.
    It allows for an optional 'database' parameter to specify the database to connect to.
    If 'database' is not provided, the connection is made without a specific database. The 'driver' parameter can
    be used to specify the database driver (default is 'SQL Server').

    Parameters:
    -----------
        server (str): The SQL Server server/instance name to establish a connection.
        user (str): The login ID.
        password (str): The password associated with the user.
        query (str): The SQL query to execute.
        database (str, optional): The name of the database to connect to (default is None).
        driver (str, optional): The database driver to use (default is 'SQL Server').

    Returns:
    --------
        pandas.DataFrame: A DataFrame containing the query result.
    
    Example:
    --------
        >>> server = '192.0.1.1\XXXXXX'
        >>> user = 'username'
        >>> password = 'pass123'
        >>> query = 'SELECT top 100 * from dventa'
        >>> database = 'pvf_prodsc'
        >>> df = query_to_df(server, user, password, query, database=database)
    """
    import pyodbc
    import pandas as pd

    if database: 
        connection_string = (
            f"DRIVER={driver};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={user};"
            f"PWD={password};"

            )

    else:
        connection_string = (
            f"DRIVER={driver};"
            f"SERVER={server};"
            f"UID={user};"
            f"PWD={password};"

            )
    connection = pyodbc.connect(connection_string)
    query = query
    df = pd.read_sql(query,connection)
    connection.close()

    return df

def previous_month_and_year():
    """
    Retrieves the month and year of the previous month.
    This function determines the year and month of the previous month based on the current date.
    It is useful for a lot of tasks, since most of the data for reports is obtained from the previous month.
    Such as the example provided, where it is required to obtain all the sales from the past month, and store
    the data on a pandas dataframe.
    

    Returns:
    --------
        tuple (m,y): A tuple containing the month (int) and year (int) of the previous month.

    Example:
    --------
        >>> month, year = previous_month_and_year()
        >>> query = f'SELECT * FROM ventas WHERE month(dventa) = {month} AND year(dventa) = {year}'
        >>> df = query_to_df(server, user, password, query)

    """
    from datetime import datetime, timedelta
    today = datetime.now()
    first_day = today.replace(day=1)
    last_day = first_day - timedelta(days=1)
    year = last_day.year
    month = last_day.month
    
    return month , year
