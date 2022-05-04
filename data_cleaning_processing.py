import pandas as pd

########
pd.set_option('display.width', 300)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
########


fpath = 'Example_Data.xlsx'
df_Data = pd.read_excel(fpath, sheet_name='Example_Data', header=1)  # head at line 1
df_DB = pd.read_excel(fpath, sheet_name='Example_DB', header=0)  # head at line 1

Data_Company_ID = df_Data.iloc[:, 0]  # [1:]tolist()
Data_Company_Name = df_Data.iloc[:, 1]


############################
### task1  Data cleaning ###
############################

# find corresponding company ID -- name set: 'map_ID'
map_ID = {}
for i in df_Data.index:
    if df_Data.loc[i, 'Company ID'] not in map_ID:
        if df_Data.loc[i, 'Company Name'] not in map_ID.values():
            map_ID[df_Data.loc[i, 'Company ID']] = df_Data.loc[i, 'Company Name']

print(map_ID)

# 1st line
df_data_cleaning = pd.DataFrame(df_Data.loc[0]).T
print(df_data_cleaning)

# 2-n lines:
for i in range(1, len(df_Data.index)):
    FY = df_Data.loc[i, 'Fiscal Year']  # Fiscal Year is the year, between 1999 and 2021
    if str(FY).isdigit():
        if int(FY) >= 1999 and int(FY) <= 2021:


            SC = df_Data.loc[i, 'SIC Code']  # SIC Code is a four-digit number
            if len(str(SC)) == 4:

                TC = df_Data.loc[i, 'Trading Currency']  # Trading Currency can only be one of two units, USD or GBP
                if TC == 'USD' or TC == 'GBP':

                    SP = df_Data.loc[i, 'SP']  # SP / CDS / APD / ARD / ADA is an integer number
                    CDS = df_Data.loc[i, 'CDS']
                    APD = df_Data.loc[i, 'APD']
                    ADA = df_Data.loc[i, 'ADA']
                    if SP >= 0 and CDS >= 0 and APD >= 0 and ADA >= 0:

                        if df_Data.loc[i, 'Company ID'] in map_ID:  # company ID matches name
                            if map_ID[df_Data.loc[i, 'Company ID']] == df_Data.loc[i, 'Company Name']:

                                df_line = pd.DataFrame(df_Data.loc[i]).T
                                df_data_cleaning = pd.concat([df_data_cleaning, df_line], ignore_index=True)


print('df_data_cleaning')
print(df_data_cleaning)


################################
### task 2  Data processing  ###
################################

df_data_processing = pd.DataFrame(columns=['Company ID', 'Company Name', 'Fiscal Year', 'Industry',
                                           'SIC Code', 'Trading Currency', 'Metric Name', 'Value'])

Metric = 'SP'
for i in df_data_cleaning.index:
    df_line = pd.DataFrame(df_data_cleaning.iloc[i][0:6]).T  # ID, name, year, industry, SIC, Currency
    value = df_data_cleaning.loc[i, Metric]
    df_metric = pd.DataFrame([[Metric, value]], columns=['Metric Name', 'Value'], index=[i])  # index=[i] to match df_line index
    df_line = pd.concat([df_line, df_metric], axis=1)  # add 'Metric Name', 'Value', horizontally
    df_data_processing = pd.concat([df_data_processing, df_line], ignore_index=True)

Metric = 'CDS'
for i in df_data_cleaning.index:
    df_line = pd.DataFrame(df_data_cleaning.iloc[i][0:6]).T  # ID, name, year, industry, SIC, Currency
    value = df_data_cleaning.loc[i, Metric]
    df_metric = pd.DataFrame([[Metric, value]], columns=['Metric Name', 'Value'], index=[i])  # index=[i] to match df_line index
    df_line = pd.concat([df_line, df_metric], axis=1)  # add 'Metric Name', 'Value', horizontally
    df_data_processing = pd.concat([df_data_processing, df_line], ignore_index=True)

Metric = 'APD'
for i in df_data_cleaning.index:
    df_line = pd.DataFrame(df_data_cleaning.iloc[i][0:6]).T  # ID, name, year, industry, SIC, Currency
    value = df_data_cleaning.loc[i, Metric]
    df_metric = pd.DataFrame([[Metric, value]], columns=['Metric Name', 'Value'], index=[i])  # index=[i] to match df_line index
    df_line = pd.concat([df_line, df_metric], axis=1)  # add 'Metric Name', 'Value', horizontally
    df_data_processing = pd.concat([df_data_processing, df_line], ignore_index=True)

Metric = 'ARD'
for i in df_data_cleaning.index:
    df_line = pd.DataFrame(df_data_cleaning.iloc[i][0:6]).T  # ID, name, year, industry, SIC, Currency
    value = df_data_cleaning.loc[i, Metric]
    df_metric = pd.DataFrame([[Metric, value]], columns=['Metric Name', 'Value'], index=[i])  # index=[i] to match df_line index
    df_line = pd.concat([df_line, df_metric], axis=1)  # add 'Metric Name', 'Value', horizontally
    df_data_processing = pd.concat([df_data_processing, df_line], ignore_index=True)

Metric = 'ADA'
for i in df_data_cleaning.index:
    df_line = pd.DataFrame(df_data_cleaning.iloc[i][0:6]).T  # ID, name, year, industry, SIC, Currency
    value = df_data_cleaning.loc[i, Metric]
    df_metric = pd.DataFrame([[Metric, value]], columns=['Metric Name', 'Value'], index=[i])  # index=[i] to match df_line index
    df_line = pd.concat([df_line, df_metric], axis=1)  # add 'Metric Name', 'Value', horizontally
    df_data_processing = pd.concat([df_data_processing, df_line], ignore_index=True)


print('df_data_processing')
print(df_data_processing)





