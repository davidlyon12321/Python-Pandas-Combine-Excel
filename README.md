Python-Pandas-Combine-Excel

Business Need:

Businesses commonly need to consolidate data from separate excel files.  Across enough files, manually copy+pasting is prohibitively time consuming.  We can use Python to pull together millions of data points from thousands of files.


Specific Case:

We receive .xlsm files from a vendor containing promotional credit amounts on their products we resell.  These are the deal-specific discount we receive, so we may receive multiple iterations of each file as we negotiate the deal with our end-customer.  Management would like a consolidated view of the credits we received on such deals in the past 12 months.


File Structure:

We have 11 .xlsm credit files (can scale up to as many files as needed) from the prior 12 months.  These cover 10 distinct sales opportunities, (Customer A received a revision, so they have a duplicate file).  We want to pull data from the most recent file for each deal.  Our file names follow a standard naming convention; 5 digit credit offer #, customer name, discount tier, and date, all separated by underscores: 24837_Customer A_Tier2_05-09-2019.xlsm.  Within each file, there are 3 tabs containing credits for different product categories: Hardware, Software, and Services.  Each tab has 1 empty row at the top, followed by a header row.  We want information from columns A, B, C, and F: Item Code, Item Description, Number of Units, and Promo Credit.


Assumptions:

No files with the same contract # AND the same date
File names have been cleaned, and are consistent
Data in the files has been cleaned
 
 
Process:

We put our input files into a subfolder.  Our python script mines the data from each file and combined it into a single file.  If there are multiple files with the same contract #, the script keeps the most recent file, removing the older duplicate data.


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
