'''
Title: Email Cleanup Script
Author: Reynaldo Cortez
Date: 2023-07-19

Description: This script cleans up an Excel file of email addresses by removing duplicates and any emails 
             that contain specified keywords. The keywords are entered by the user when the script is run.
'''

import pandas as pd
import numpy as np

# Get the file path and keywords from the user
file_path = input("Enter the path to your Excel file: ")
keywords = input("Enter the keywords to search for, separated by commas: ").split(',')

# Read the Excel file
df = pd.read_excel(file_path)

# Remove duplicates
df = df.drop_duplicates(subset=['Emails'])

# Remove rows containing the keywords
for keyword in keywords:
    df = df[np.logical_not(df['Emails'].str.contains(keyword.strip(), case=False))]

# Save the cleaned DataFrame back to the original Excel file
df.to_excel(file_path, index=False)

print("Cleanup completed successfully!")
