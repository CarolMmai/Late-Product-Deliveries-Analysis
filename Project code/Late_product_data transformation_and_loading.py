# Load to SQL for Late_Product_deliveries report---------------------------------------------------------------------------------------------------------------------------------------------------

import os
import pandas as pd
from sqlalchemy import create_engine
import datetime
import re
import time

# Define the folder path
folder_path = r"E:\Transporter_late_product_deliveries_data\Files_to_append"
log_directory = r"E:\Transporter_late_product_deliveries_data\Log_Files"

# List all files in the folder with the prefix "FF3X" and extension ".xlsx"
file_list = [f for f in os.listdir(folder_path) if f.startswith("FF3") and f.endswith(".xlsx")]

# SQL Server connection details
conn_str = (
    r'Trusted_Connection=yes;'
    r'Driver={ODBC Driver 17 for SQL Server};'
    r'Server=SERVER54321;'
    r'Database=ProductDeliveriesInfo;'
)

# Get current date and time for log file
current_datetime = datetime.datetime.now()
log_filename = f"db_populate_{current_datetime.strftime('%Y%m%d_%H%M%S')}.txt"
log_path = os.path.join(log_directory, log_filename)


# Create an empty DataFrame to store the appended data
appended_data = []

# Initialize a variable to store the total number of rows
total_rows = 0

# Loop through each Excel file
for file in file_list:
    if file.endswith(".xlsx") and file.startswith("FF3"):
        file_path = os.path.join(folder_path, file)
    try:
        # Read Excel files
        df = pd.read_excel(file_path)
        
        # Calculate the number of rows in the DataFrame
        num_rows = len(df)
            
        # Add the number of rows to the total
        total_rows += num_rows
        
        appended_data.append(df)
        
        combined_df = pd.concat(appended_data, ignore_index=True)
        
        # Remove characters not allowed in SQL Server headers
        invalid_chars = r'[-,#:.,/\\?@& ]'
        combined_df.columns = [re.sub(invalid_chars, "_", col) for col in combined_df.columns]

        # Add a column named "upload DTM" with current date and time
        upload_dtm = datetime.datetime.now()
        combined_df["upload DTM"] = upload_dtm

        # Create the SQLAlchemy engine
        engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

        # Write the DataFrame to SQL Server, replacing the existing table
        ProductDeliveries_table_name = "ProductDeliveries"
        combined_df.to_sql(ProductDeliveries_table_name, con=engine, if_exists='replace', index=False)
      
      
        # Execute the stored procedure after table creation
        exec_sp_query = "EXEC [dbo].[ProductDeliveries_Update_&_Transform] @ProductDeliveriesInfo = 'ProductDeliveriesInfo', @ProductDeliveries = 'ProductDeliveries', @ProdDel_Transformed_Tbls = 'ProdDel_Transformed_Tbls', @ProdDel_Bronze = 'ProdDel_Bronze';"
      
        # Calculate the number of rows in the combined DataFrame
        num_rows = len(combined_df)
        
        # Write log information
        with open(log_path, "a") as log_file:
            log_message = f"Table '{ProductDeliveries_table_name}' has been created in the database 'ProductDeliveriesInfo' using 'replace' mode.\n"
            log_file.write(log_message)
            log_message = f"File '{file}' has been successfully uploaded to SQL database.\n"
            log_file.write(log_message)
            log_message = f"Number of rows counted are: {num_rows}\n"
            log_file.write(log_message)
         
        # Calculate the total number of rows in the appended data
        total_num_rows = len(combined_df)
            
        # Write a line in the log file with the total number of rows
        with open(log_path, "a") as log_file:
            log_message = f"Total number of rows from all files: {total_num_rows}\n"
            log_file.write(log_message)
        
        
        # Delete the uploaded file
        os.remove(file_path)
        with open(log_path, "a") as log_file:
            log_message = f"File '{file}' has been successfully deleted.\n"
            log_file.write(log_message)
    except Exception as e:
        # Handle exceptions and log errors
        with open(log_path, "a") as log_file:
            error_message = f"Error processing file '{file}': {str(e)}\n"
            log_file.write(error_message)

print(f"Log file '{log_filename}' has been created in the directory '{log_directory}'.")