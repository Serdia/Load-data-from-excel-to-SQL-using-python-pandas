# iterating through each excel sheet
# creating list of dataframes
# joining list of dataframes with each other on Division and BudgetDate columns

import pandas as pd
import datetime

sheet_names = ['Gross Revenue','Gross Margin','EBITDA','Commission Income','Commission Expense','Fee Income']
list_of_df = []
for sheet in sheet_names:
    df = pd.read_excel(r'\\username\Documents\Python\Projects\Excel to SQL DB\AFH_tblBudget_2019.xlsx', sheet_name = sheet, index=False)
    df = pd.melt(df,id_vars=['Division'], var_name ='BudgetDate', value_name = sheet)
    list_of_df.append(df)


full_df = list_of_df[0].merge(list_of_df[1],on=['Division','BudgetDate']).merge(list_of_df[2],on=['Division','BudgetDate']).merge(list_of_df[3],on=['Division','BudgetDate']).merge(list_of_df[4],on=['Division','BudgetDate']).merge(list_of_df[5],on=['Division','BudgetDate'])
full_df['BudgetDate']= pd.to_datetime(full_df['BudgetDate']) 
full_df = full_df.sort_values(['Division','BudgetDate'])

#replacing nan values to 0
full_df.fillna(0,inplace=True)
full_df = full_df.reset_index(drop=True)
#print(full_df)

# this part inserts dataframe into sql database (mejames)

conn = pyodbc.connect('DRIVER={SQL Server};server=10.26.8.18;DATABASE=MEJAMES;Trusted_Connection=yes;')
cursor = conn.cursor()

for index,row in full_df.iterrows():
    cursor.execute("""INSERT INTO dbo.AFH_tblBudget_2019
                   (
                         Division
                        ,BudgetDate
                        ,GrossRevenue
                        ,GrossMargin
                        ,EBITDA
                        ,CommissionIncome
                        ,CommissionExpense
                        ,FeeIncome
                  ) 
                         values (?,?,?,?,?,?,?,?)""", row['Division'], 
                                                   row['BudgetDate'], 
                                                   row['Gross Revenue'],
                                                    row['Gross Margin'] ,
                                                    row['EBITDA'] ,
                                                    row['Commission Income'],
                                                    row['Commission Expense'],
                                                    row['Fee Income'])
    
    conn.commit()
cursor.close()
conn.close()   

