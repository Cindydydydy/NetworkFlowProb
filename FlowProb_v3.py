import pandas as pd
from pandas import DataFrame
import numpy as np

'''
Create new .xlsx file
'''
# Create an empty DataFrame
df_out = pd.DataFrame()

# Save the DataFrame to a .xlsx file
output_file = 'output_data_5.xlsx'

'''
Find Delivery & Demand & insert into new .xlsx file named as output_file
'''
# Read the .xlsx file
file = 'NetworkFlowProblem-Data.xlsx'
df = pd.read_excel(file, sheet_name='Input5') #!!! change input sheet name here

# Find the certain content in the 'for_process' column
result_delivery = df.loc[df['for_process'] == 'Delivery']
result_delivery.sort_values(by=['Amount'],axis=0,ascending=False,inplace=True)

rows_delivery = [row for row in result_delivery.iterrows()]


cols5 = ["Sent cnt5",      "Process5",    "Cnt5",             "Week5", "Amount5"]
cols =  ["send_from_cnt", "for_process", "to_processing_cnt", "Week" , "Amount"]
# for row in result_delivery.iterrows():
for i in range(len(cols)):
   df_out.insert(i, cols5[i], result_delivery[cols[i]], allow_duplicates=False)
df_out.insert(len(cols), 'Demand', np.arange(1,len(rows_delivery)+1),  allow_duplicates=True)

df_out.to_excel(output_file, index=False, engine='openpyxl')

# obtain the length of df_out 
rows_df_out = [row for row in df_out.iterrows()]
num_demands = len(rows_df_out)

'''
Generation function: cols
'''
def geneFunc(val): # val 4,3,2,1
    cols_val = []
    cols_val.append("Sent cnt" + str(val))
    cols_val.append("Process" + str(val))
    cols_val.append("Cnt" + str(val))
    cols_val.append("Week" + str(val))
    cols_val.append("Amount" + str(val))
    return cols_val


'''
Insertion function
'''
def insertProcess(val, cols,cols_val,process, rows_df_out, result_delivery, df_out): # process \in str
    
    sent = 'Sent cnt' + str(val+1)
    cols_sentCnt = []
    cols_Process = []
    cols_Cnt = []
    cols_Week = []
    cols_Amount = []
    
    # Find the certain content in the 'for_process' column
    result_process = df.loc[df['for_process'] == process]
    result_process.sort_values(by=['Amount'],axis=0,ascending=False,inplace=True)
    
    for j in range(len(rows_df_out)):
        if process == 'Forwarding': 
            pos_process = result_process.loc[result_process["to_processing_cnt"] == df_out[sent][rows_df_out[j][0]]]
            pos_process_amount = pos_process.loc[ pos_process["Amount"] >= result_delivery["Amount"][rows_delivery[j][0]]]
        else:
            pos_process = result_process.loc[result_process["to_processing_cnt"] == df_out[sent][rows_df_out[j][0]]]
            pos_process_amount = pos_process.loc[ pos_process["Amount"] >= df_out["Amount5"][rows_df_out[j][0]]]
        
        if pos_process_amount is not None:
            idx = [row for row in pos_process_amount.iterrows()][0][0] # 1st []: remove list; 2nd []: extract out first element  
            cols_sentCnt.append(pos_process_amount.loc[idx]['send_from_cnt'])
            cols_Process.append(pos_process_amount.loc[idx]['for_process'])
            cols_Cnt.append(pos_process_amount.loc[idx]['to_processing_cnt'])
            cols_Week.append(pos_process_amount.loc[idx]['Week'])
            if process == 'Forwarding': 
                cols_Amount.append(result_delivery["Amount"][rows_delivery[j][0]])
            else:
                cols_Amount.append(df_out["Amount5"][rows_df_out[j][0]])
            
            # update amount valus in the "forwarding" amount part
            diff = pos_process_amount.loc[idx,'Amount'] - result_delivery["Amount"][rows_delivery[j][0]] 
            df.loc[idx, 'Amount'] = diff
            pos_process_amount.loc[idx,'Amount'] = df.loc[idx, 'Amount']
            pos_process.loc[idx, 'Amount'] = df.loc[idx, 'Amount']
            result_process.loc[idx, 'Amount'] = df.loc[idx, 'Amount']
            
    dict_cols = {cols_val[0]:cols_sentCnt, cols_val[1]: cols_Process, cols_val[2]: cols_Cnt, cols_val[3]: cols_Week, cols_val[4]: cols_Amount}
    for i in range(len(cols)):
        df_out.insert(i, cols_val[i], dict_cols[cols_val[i]], allow_duplicates=True)
    return df_out

'''
Forwarding & Treatment & Conditioning & Sourcing   
'''
remainProcess = ['Sourcing', 'Conditioning', 'Treatment', 'Forwarding'] #['Forwarding', 'Treatment', 'Conditioning', 'Sourcing']
for val in range(len(remainProcess),0,-1): 
    cols_val = geneFunc(val)
    process = remainProcess[val-1]
    df_out = insertProcess(val, cols,cols_val,process, rows_df_out, result_delivery, df_out)
    df_out.to_excel(output_file, index=False, engine='openpyxl')
