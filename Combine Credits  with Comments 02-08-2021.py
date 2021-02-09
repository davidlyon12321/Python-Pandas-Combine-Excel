#Make sure you have Python 3.7, and the most recent verisons of pip, pandas, numpy, openpyxl, xlrd

import numpy as np
import pandas as pd
import glob
import os
import re

df = pd.DataFrame()

os.chdir('<INPUT FOLDER>')

for f in glob.glob("*.xlsx"):
	#Split the filename into Contract #, Customer Name, Discount Tier Level, and Discount Date
	value = f.split('_')
	contract = value[0]
	customer = value[1]
	level = value[2]
	initialissuedate = value[3].partition('.')
	issuedate = initialissuedate[0]
	#Pull data from each tab, then add the Customer, Contract, Discount Level, and Discount Date to each row
	data = pd.concat([pd.read_excel(f, sheet, header=1, usecols="A,B,C,F").assign(Customer=customer).assign(Contract=contract).assign(Level=level).assign(Issuedate=issuedate) for sheet in ['Hardware', 'Software', 'Services']], sort=False)
	df = df.append(data, sort=False, ignore_index=True)
	#Drop all rows where quantity = 0
	df = df[df['Number of Units'] !=0]
	#Drop all rows with blank/missing/null data
	df.replace('', np.nan, inplace=True)
	df.dropna(axis=0,how='any',inplace=True)

#Point to where we want our output file to go
os.chdir('<OUTPUT FOLDER>')
#Take only the rows from the most recent offer for each customer
df['Issuedate'] = pd.to_datetime(df['Issuedate'])
idx = df.groupby(['Contract'])['Issuedate'].transform(max) == df['Issuedate']
#Output the final dataset to a .csv
df[idx].to_csv('output1.csv')