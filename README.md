Python-Pandas-Combine-Excel

Businesses commonly need to consolidate data from separate excel files. Across enough files, manually copy+pasting is prohibitively time consuming. We can use Python to pull together millions of data points from thousands of files.

Business Case:

We receive .xlsm files from a vendor containing promotional credit amounts on their products we resell. These are the deal-specific discount we receive, so we may receive multiple iterations of each file as we negotiate the deal with our end-customer. Management would like a consolidated view of the credits we received on such deals in the past 12 months.

Input Files:

Files have a standard naming convention, separated by underscores

Some contracts have multiple files; we only want the most recent file for each contract

No contract received multiple credit files in the same day

Filenames and content have been cleaned


File Structure:

Our files have 4 desirable columns: A,B,C,F

One empty row at the top

Row #2 is the header row

3 Tabs: Hardware, Software, and Services

We don’t want rows where Number of Units = 0


Python Script:

Use Pandas to pull the data into a dataframe

Add data from each file in the inputs folder

If there is already data for a contract, only keep data from most recent file

Add the contract #, customer name, discount tier, and file date to each row for easy analysis


Output:

Shows the desired columns

Excludes where Number of Units = 0

Includes only the most recent file’s data for each contract



Files:

Combine Credits xx-xx-xxxx.py: Python script w/o notes

Combine Credits with Comments xx-xx-xxxx.py: Python script with notes

Input files:

12310_Customer D_T0_01-22-2020.xlsx

19020_Customer I_T0_06-06-2020.xlsx

22301_Customer J_T2_10-25-2020.xlsx

23151_Customer C_T1_04-11-2020.xlsx

33412_Customer F_T0_09-19-2020.xlsx

45829_Customer A_T1_02-22-2020.xlsx

45829_Customer A_T1_05-22-2020.xlsx

48120_Customer B_T1_11-14-2020.xlsx

57329_Customer H_T3_02-24-2020.xlsx

59201_Customer E_T2_08-15-2020.xlsx

89210_Customer G_T0_07-21-2020.xlsx
