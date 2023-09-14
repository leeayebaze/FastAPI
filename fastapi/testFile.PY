import warnings
from fastapi.responses import JSONResponse
import pandas as pd
import pyodbc
import os
import glob
import openpyxl
import logging
from fastapi import FastAPI
# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

# Configuration parameters (customize these based on your environment)
CONFIG = {
    'batch_number': '2346',
    'server': '10.60.41.10,4120',
    'database': 'BusinessIntelligence',
    'username': 'dataanalyst',
    'password': 'Waeg@gah5y',
    'excel_path': r'C:\Users\DELL\Desktop\Marvin',
    'file_pattern': 'Batch {batch_number}.xlsxx'
}

def establish_db_connection(server, database, username, password):
    try:
        connection_string = f'DRIVER=SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password};TrustServerCertificate=yes'
        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            return conn, cursor
    except Exception as e:
        logging.error(f"Error establishing a database connection: {str(e)}")
        return None, None

def fetch_data(cursor, batch_number):
    try:
        query = """
            SELECT DATE_TIME, TRN_REF, BATCH, TXN_TYPE, ISSUER, ACQUIRER, AMOUNT, FEE, BENEFICIARY_ENTITY, ABC_COMMISSION,
            CASE WHEN RESPONSE_CODE = 'null' THEN 'Failed' 
            WHEN TRY_CAST(RESPONSE_CODE AS INT) = 0 THEN 'Successful' ELSE 'Failed' END AS TRAN_STATUS
            FROM Transactions WHERE (RESPONSE_CODE IS NOT NULL OR RESPONSE_CODE IN ('0')) AND BATCH = ? AND ISSUER_CODE != '730147'
            AND TXN_TYPE NOT IN ('ACI', 'AGENTFLOATINQ', 'MINI')
            """
        cursor.execute(query, batch_number)
        datafile = pd.DataFrame(cursor.fetchall(), columns=[desc[0] for desc in cursor.description])
        return datafile
    except Exception as e:
        logging.error(f"Error fetching data from the database: {str(e)}")
        return None

def read_excel_file(file_path, sheet_name):
    try:
        with pd.ExcelFile(file_path) as xlsx:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, usecols=[0, 1, 2, 7, 8, 9, 11], skiprows=0)
        return df
    except Exception as e:
        logging.error(f"An error occurred while opening the Excel file: {e}")
        return None

def manipulate_data(df):
    try:
        # Convert 'DATE_TIME' column to datetime
        df['DATE_TIME'] = pd.to_datetime(df['DATE_TIME'])
        df['DATE_TIME'] = df['DATE_TIME'].dt.strftime('%d%m%Y').str.replace(r'[-/\\.]', '')
        df = df[['DATE_TIME', 'TRN_REF', 'BATCH', 'TXN_TYPE', 'AMOUNT', 'FEE', 'ABC_COMMISSION']]

        # Specify the columns you want to convert to strings
        columns_to_convert = ['DATE_TIME', 'TRN_REF', 'BATCH']

        # Convert the specified columns to strings
        for column in columns_to_convert:
            df[column] = df[column].astype(str)

        # Define a function to remove the first zero from the left for strings greater than 11 characters
        def remove_first_zero(input_str):
            if len(input_str) > 11 and input_str.startswith('0'):
                # Remove the first zero from the left
                cleaned_str = input_str[1:]
                return cleaned_str
            else:
                # If the string is 11 or fewer characters or doesn't start with zero, return it unchanged
                return input_str

        # Apply the function to the 'TRN_REF' column
        df['TRN_REF'] = df['TRN_REF'].apply(remove_first_zero)

        # Create a concatenated column for the dataset
        df['Concat'] = (
            df['TRN_REF'].astype(str) +
            df['DATE_TIME'].astype(str) +
            df['BATCH'].astype(str))

        return df
    except Exception as e:
        logging.error(f"An error occurred while manipulating data: {str(e)}")
        return None

app = FastAPI()



# Refactor the main function to accept a batch_number and return the result
def main(batch_number: str):
    try:
        conn, cursor = establish_db_connection(CONFIG['server'], CONFIG['database'], CONFIG['username'], CONFIG['password'])
        CONFIG['batch_number'] = batch_number
        if conn and cursor:
            datafile = fetch_data(cursor, CONFIG['batch_number'])
            cursor.close()
            conn.close()

            if datafile is not None:
                # Data manipulation for 'datafile'
                datafile = manipulate_data(datafile)

                # Create a pattern to match the specific value in the filename
                file_pattern = f'Batch {CONFIG["batch_number"]}.xlsxx'

                # Get a list of all Excel files in the folder that match the pattern
                excel_files = glob.glob(os.path.join(CONFIG['excel_path'], file_pattern))

                # Check if the matching file was found
                if not excel_files:
                    logging.error(f"No matching Excel file found for '{file_pattern}'.")
                else:
                    # Assuming there's only one matching file, you can open it
                    matching_file = excel_files[0]

                    # Read the Excel file using pandas with a context manager
                    SABSfile_ = read_excel_file(matching_file, 'Transaction Report')

                    if SABSfile_ is not None:
                        # Data manipulation for 'SABSfile_'
                        SABSfile_ = manipulate_data(SABSfile_)

                        # More processing for SABSfile_
                        ##Remove Time from Date
                        SABSfile_['DATE'] = SABSfile_['DATE'].dt.strftime('%d%m%Y').str.replace(r'[-/\\.]', '')

                        # Specify the columns you want to convert to strings
                        columns_to_convert = ['DATE', 'Trx Ref.', 'Batch Ref']

                        # Convert the specified columns to strings
                        for column in columns_to_convert:
                            SABSfile_[column] = SABSfile_[column].astype(str)

                        # Define a function to pad strings with zeros
                        def pad_strings_with_zeros(input_str):
                            if len(input_str) < 12:
                                # Calculate the number of zeros needed
                                num_zeros = 11 - len(input_str)
                                # Add zeros to the beginning of the string
                                padded_str = '0' * num_zeros + input_str
                                return padded_str
                            else:
                                # If the string is already 11 or more characters, return it unchanged
                                return input_str

                        # Apply the function to the 'Trx Ref.' column
                        SABSfile_['Trx Ref.'] = SABSfile_['Trx Ref.'].apply(pad_strings_with_zeros)

                        # Create a concatenated column for the SABS dataset
                        SABSfile_['Concat_sabs'] = (
                            SABSfile_['Trx Ref.'].astype(str) +
                            SABSfile_['DATE'].astype(str) +
                            SABSfile_['Batch Ref'].astype(str))

                        # Merge the two DataFrames on the 'Concat_sabs' and 'Concat_db' columns
                        merged_data = pd.merge(SABSfile_, datafile, left_on='Concat_sabs', right_on='Concat_db', how='outer')

                        # Calculate the difference between Amounts and Commission Amounts
                        merged_data['Amount Var'] = merged_data['Amount'] - merged_data['AMOUNT']
                        merged_data['Comm Var'] = merged_data['ABC Commission'] - merged_data['ABC_COMMISSION']

                        # Convert 'Date' column back to datetime format if needed
                        # merged_data['DATE'] = pd.to_datetime(merged_data['DATE'], format='%d%m%Y')

                        # Filter rows that have differences
                        filtered_df = merged_data[(merged_data['Amount Var'] < 0) | (merged_data['Comm Var'] < 0)]

                        # Select the desired columns for the final output
                        filtered_df = filtered_df[['DATE', 'Trx Ref.', 'TRN_REF', 'Batch Ref', 'Amount', 'Amount Var', 'Comm Var']]
                        # Return the resulting DataFrame
        return filtered_df

    except Exception as e:
        logging.error(f"An error occurred in the main process: {str(e)}")
   

@app.get("/transactions/")
async def get_transactions(batch_number: str):
    result = main(batch_number)
    
    if result is None:
        return JSONResponse(content={"error": "Error processing request."}, status_code=500)

    # Convert DataFrame to dictionary for JSON response
    return result.to_dict(orient="records")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)