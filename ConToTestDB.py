import pyodbc 
import pandas as pd
import numpy as np
from pandasql import sqldf

# Set up your target variables
server = 'localhost'
database = 'Test'

# Create a connection to the database
conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server}; \
                      Server=' + server + '; \
                      Database=' + database + '; \
                      Trusted_Connection=yes;')

# Create connection cursor            
cursor = conn.cursor()
cursor.execute('DELETE FROM TrackSamples')

# Prepare the insert query
insert_query = '''INSERT INTO TrackSamples (SampNo, ResidenceID, FileName, SamplingProgram)
                  VALUES (?, ?, ?, ?);'''


# Grab the excel file needed
sample_certs = pd.ExcelFile(r'C:\Users\kelse\OneDrive\Documents\DummyWorkbook.xlsx')

# put the contents of the file into a data frame
df = pd.read_excel(sample_certs, usecols=[0, 1, 2, 3])

print(df)
# pysqldf = lambda q: sqldf(q, globals())

# Insert every row from the dataframe into the database
for index, row in df.iterrows():
    values = (row[0], row[1], row[2], row[3])
    cursor.execute(insert_query, values)













########################PYODBC##########################

#data for the database
person_data = [['Kelsey', 'Bonnicksen', '02/06/1995', 'Female' ],
                ['Andrew', 'Wheeler', '03/03/1993', 'male']]

# for driver in pyodbc.drivers():
#     print(driver)




# Define our insert query


# loop through all data

    

# commit the inserts
conn.commit()

cursor.execute('SELECT * FROM Person')

# for row in cursor:
#     print(row)

# close cursor and connection
cursor.close()
conn.close()

# cursor.execute('SELECT * FROM Test.People')

# for row in cursor:
#     print(row)
#     print('Something');

#     print('I hope this will work');

# print('I hope this will work')
# print('OH MY GOD IT WORKED')
# exit()