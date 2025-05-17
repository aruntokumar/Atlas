import pandas as pd
from datetime import datetime, timedelta
import getpass

user_name = getpass.getuser()

print(" Author : ArunKumar & Gokulakannan ")

# Read the Excel file into a pandas DataFrame
df = pd.read_excel('C:\\Users\\'+user_name+'\\Desktop\\inflow_Asin\\inflow.xlsx')

# Convert column B to datetime if it's not already in datetime format
df['creation_date'] = pd.to_datetime(df['creation_date']).dt.date


# Filter data for the last 7 days
today = datetime.now()
today = today.replace(hour=00, minute=00, second=00,microsecond=00)

start_date = today - timedelta(days=8)
end_date = today - timedelta(2)

print("Filter Stage 1 : " + str(start_date.date())+" to "+str(end_date.date())+" new_unreleased_child_plan is Y")

filtered_data_1 = df[(df['creation_date'] >= start_date.date()) & (df['creation_date'] <= end_date.date()) & (df['unreleased_parent_plan_in_wait_time']=='Y')]


print("Filter Stage 2 : " + str(start_date.date())+" to "+str(end_date.date())+" unreleased_parent_plan_in_wait_time is Y")

filtered_data_2 = df[(df['creation_date'] >= start_date.date()) & (df['creation_date'] <= end_date.date()) & (df['new_unreleased_child_plan']=='Y')]

final_data = pd.concat([filtered_data_1,filtered_data_2])

# Find and replace
col_A = {'Other': 'EU_MVR', 'EU MVR': 'EU_MVR', 'Retail Business':'RETAIL_BUSINESS'}
col_E = {193:'gl_apparel',309:'gl_shoes',200:'gl_sports',198:'gl_luggage'}
col_D = {4:'DE',5:'FR',3:'UK',44551:'ES',35691:'IT'}

# replace values using the .map() method
final_data['org_owner'] = final_data['org_owner'].map(col_A).fillna(df['org_owner'])
final_data['marketplace_id'] = final_data['marketplace_id'].map(col_D).fillna(df['marketplace_id'])
final_data['gl_product_group'] = final_data['gl_product_group'].map(col_E).fillna(df['gl_product_group'])

#Changing Heading name
final_data = final_data.rename(columns={'org_owner':'BUS','gl_product_group':'gl','marketplace_id': 'mp', 'commitment_id': 'C-plan','creation_date':'crimson_creation_date'})

# Inserting new columns
final_data['Actual_VC'] =''
final_data['cmbs_creation_date'] = final_data['crimson_creation_date']
final_data['source_submission_ids']=final_data['brand']=final_data['vendor_code']

# Final Data ready
final_df = final_data[['BUS','gl','source_submission_ids','vendor_code','brand','mp','C-plan','plan_name','crimson_creation_date','cmbs_creation_date','vendor_manager','status','Actual_VC']]

final_df.to_excel('C:\\Users\\'+user_name+'\\Desktop\\inflow_Asin\\Output.xlsx',index=False)

print("Successfully Completed !!")