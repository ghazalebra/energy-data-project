import pandas as pd
import numpy as np
from datetime import date

num_cols = 29

input_column_names = ["Date", "Time", "OA RH", "OA TEMP", "CH-1 \n冰水主機", "PCHP-1 \n冰水泵", "CDWP-1 \n冷卻水泵",
                        "CH-2 \n冰水主機", "PCHP-2\n冰水泵", "CDWP-2 \n冷卻水泵", "CH-3 \n冰水主機", "PCHP-3\n冰水泵",
                        "CDWP-3\n冷卻水泵", "CT-1 \n冷卻水塔", "CT-2 \n冷卻水塔", "CT-3 \n冷卻水塔", "總耗電量", "總冰水流量\n(LPM)",
                        "冰水回水溫度(。C)", "冰水出水溫度(。C)", "總冷凍能力\n(RT)", "總冷卻水流量(LPM)", "冷卻水回水溫度(。C)",
                        "冷卻水出水溫度(。C)", "總冷卻能力(RT)", "冰水主機耗電量(KW)", "冰水主機耗電量(RT)", "熱平衡%", "熱平橫>15%"]


output_columns_names = [
    "Date",
    "每日平均外氣溫度(。C)",
    "每日平均外氣濕度(%RH)",
    "CH-1 冰水主機累樍耗電量(kWh)",
    "PCHP-1 冰水泵累樍耗電量(kWh)",
    "CDWP-1 冷卻水泵累樍耗電量(kWh)",
    "CH-2 冰水主機累樍耗電量(kWh)",
    "PCHP-2 冰水泵累樍耗電量(kWh)",
    "CDWP-2 冷卻水泵累樍耗電量(kWh)",
    "CH-3 冰水主機累樍耗電量(kWh)",
    "PCHP-3 冰水泵累樍耗電量(kWh)",
    "CDWP-3 冷卻水泵累樍耗電量(kWh)",
    "CT-1 冷卻水塔累樍耗電量(kWh)",
    "CT-2 冷卻水塔累樍耗電量(kWh)",
    "CT-3 冷卻水塔累樍耗電量(kWh)",
    "總累樍耗電量(kWh)",
    "系統累積製冷能力(RT-H)"
]

output_column_names2 = ["冷凍 能力", 
                        "冰水機 耗電量", 
                        "冰水泵 耗電量", 
                        "冷卻水泵 耗電量", 
                        "冷卻水塔 耗電量", 
                        "系統 效率值", 
                        "約定冷能 需求量", 
                        "改善後 能源耗用量", 
                        "改善後 能源費用"
                        ]

table2_units = ["RTh", "kWh", "kWh", "kWh",	"kWh", "kW/RT", "RT-h/年", "kWh/年", "元/年"]

output_columns_names3 = [
    "紀錄天數",
    "應具備資料數",
    "有效數據筆數",
    "無效數具比例",
    "熱平橫>15%筆數",
    "熱平橫>15%比例",
    "熱平橫<15%比例",
    "耗能指標值KW/RT"
]

file_path = 'input-data.xls'
try:
    df = pd.read_excel(file_path, skiprows=[0], dtype={'OA TEMP': float})
    input_table = df.iloc[:, :num_cols]
    print(input_table)
except Exception as e:
    print("An error occurred while reding the data:", e)

start_date = pd.to_datetime(input('Enter the start date please: ') or '2023-8-1')
end_date = pd.to_datetime(input('Enter the end date please: ') or '2023-8-31')
target_dates = pd.date_range(start_date, end_date, freq='d').tolist()
# print(target_dates)

output_table1 = pd.DataFrame(columns=output_columns_names)
output_table2 = pd.DataFrame(columns=output_column_names2)
output_table3 = pd.DataFrame(columns=output_columns_names3)

def select_rows_by_date(target_date):
    try:
        input_table['Date'] = pd.to_datetime(input_table['Date'])
        dates = input_table["Date"]
        
        # Select rows with the target data
        target_rows = input_table[dates == pd.to_datetime(target_date)]
        
        return target_rows
    
    except Exception as e:
        print("An error occurred:", e)
        return None
    
def create_table1(input_rows):
    if input_rows is not None:
        dates = input_rows.keys()
        for i, date in enumerate(dates):

            data = input_rows[date]
            output_table1.loc[i] = [date, data['OA TEMP'].mean(), data['OA RH'].mean()] + [np.round(data[input_column_names[j]].sum()/60) for j in range(4, 17)] + [np.round(data[input_column_names[20]].sum()/60)]

        output_table1.iloc[:, 1:] = output_table1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
        # print(output_table1)
        output_table1.loc[len(dates)] = ['合計'] +  [np.round(output_table1.loc[:, output_columns_names[j]].mean(), 2) for j in range(1, 3)] + [output_table1.loc[:, output_columns_names[j]].sum() for j in range(3, 17)]
        output_table1.iloc[len(dates), 1:] = output_table1.iloc[len(dates), 1:].apply(pd.to_numeric, errors='coerce')

    else:
        print("Failed to read the Excel file or date not found.")

def create_table2():
    if output_table1 is not None:
        last_row = len(output_table1.index)
        # columns containing CH
        column_groups = {'CH': [], 'PCHP': [], 'CDWP': [], 'CT': []}

        for column_name in output_columns_names:
            if 'PCHP' in column_name:
                column_groups['CH'].append(column_name)
            elif 'CH' in column_name:
                column_groups['PCHP'].append(column_name)
            elif 'CDWP' in column_name:
                column_groups['CDWP'].append(column_name)
            elif 'CT' in column_name:
                column_groups['CT'].append(column_name)

        system_efficiency_value = np.round(output_table1.at[last_row-1, '總累樍耗電量(kWh)'] / output_table1.at[last_row-1, '系統累積製冷能力(RT-H)'], 3)
        agreed_cold_energy_demand = 2203200
        improved_energy_consumption = system_efficiency_value*agreed_cold_energy_demand
        improved_energy_cost = improved_energy_consumption*2.6
        new_row = [output_table1.at[last_row-1, '系統累積製冷能力(RT-H)']] + [output_table1.loc[last_row-1, column_groups[group]].sum() for group in column_groups] + [system_efficiency_value, agreed_cold_energy_demand, improved_energy_consumption, improved_energy_cost]
        output_table2.loc[1] = new_row
    else:
        print("Failed to read the Excel file or date not found.")

def create_table3():
    total_data = input_table.shape[0]
    hot_flat_horizontal_15_number_of_transactions = input_table.loc[:, input_column_names[-1]].sum()
    output_table3.loc[1] = [len(target_dates), total_data, total_data, (total_data - total_data)/(1440*31), hot_flat_horizontal_15_number_of_transactions, hot_flat_horizontal_15_number_of_transactions/(1440*31), 1 - hot_flat_horizontal_15_number_of_transactions/(1440*31), output_table2.loc[1, "系統 效率值"]]
    
print('Selecting the rows...')
input_rows = {target_date: select_rows_by_date(target_date) for target_date in target_dates}
print('Creating table 1...')
create_table1(input_rows)
print('Creating table 2...')
create_table2()
create_table3()

print(output_table1)
print(output_table2)
print(output_table3)


