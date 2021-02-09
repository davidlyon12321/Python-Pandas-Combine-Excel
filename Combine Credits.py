import numpy as np
import pandas as pd
import glob
import os
import re
df = pd.DataFrame()
os.chdir('C:/Users/Rainbird/Documents/Python/Promo Credits/Credit Files')
for f in glob.glob("*.xlsx"):
	value = f.split('_')
	contract = value[0]
	customer = value[1]
	level = value[2]
	initialissuedate = value[3].partition('.')
	issuedate = initialissuedate[0]
	data = pd.concat([pd.read_excel(f, sheet, header=1, usecols="A,B,C,F").assign(Customer=customer).assign(Contract=contract).assign(Level=level).assign(Issuedate=issuedate) for sheet in ['Hardware', 'Software', 'Services']], sort=False)
	df = df.append(data, sort=False, ignore_index=True)
	df = df[df['Number of Units'] !=0]
	df.replace('', np.nan, inplace=True)
	df.dropna(axis=0,how='any',inplace=True)
os.chdir('C:/Users/Rainbird/Documents/Python/Promo Credits/Output')
df['Issuedate'] = pd.to_datetime(df['Issuedate'])
idx = df.groupby(['Contract'])['Issuedate'].transform(max) == df['Issuedate']
df[idx].to_csv('output1.csv')