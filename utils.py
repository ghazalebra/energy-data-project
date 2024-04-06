import pandas as pd
import numpy as np



class Data:

    def __init__(self, file=None, start_date='2023-08-01', start_time="0:00:00", end_date='2023-08-31', end_time="23:59:00", output_file_name="output.xlsx"):
        self.file = file
        self.start_date = pd.to_datetime(start_date) if start_date else '2023-08-01' 
        self.end_date = pd.to_datetime(end_date) if end_date else '2023-08-31' 
        self.output_file_name = output_file_name

        self.start_time =  pd.to_datetime(start_time).time() if start_time else "0:00:00"
        self.end_time = pd.to_datetime(end_time).time() if end_time else "23:59:00"

        self.target_dates = pd.date_range(start=self.start_date, end=self.end_date)

        self.num_days = len(self.target_dates)
        
    
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
    output_column_names2 = ["冷凍能力(RTh)", 
                            "冰水機耗電量(kWh)", 
                            "冰水泵耗電量(kWh)", 
                            "冷卻水泵耗電量(kWh)", 
                            "冷卻水塔耗電量(kWh)", 
                            "系統效率值(kW/RT)", 
                            "約定冷能需求量(RT-h/年)", 
                            "改善後能源耗用量(kWh/年)", 
                            "改善後能源費用(元/年)"
                            ]
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

    output_table1 = pd.DataFrame(columns=output_columns_names)
    output_table2 = pd.DataFrame(columns=output_column_names2)
    output_table3 = pd.DataFrame(columns=output_columns_names3)

    def extract_input_table(self):
        try:
            df = pd.read_excel(self.file, skiprows=[0], dtype={'OA TEMP': float})
            self.input_table = df.iloc[:, :self.num_cols]
            # self.input_table['Date'] = [str(self.input_table['Date'][i].date()) for i in range(len(self.input_table['Date']))]
            # self.input_table['Time'] = [str(self.input_table['Time'][i]) for i in range(len(self.input_table['Time']))]

            # self.input_table['Datetime'] = pd.to_datetime(self.input_table['Date'] + ' ' + self.input_table['Time'])
            # self.input_table.set_index('Datetime', inplace=True)
        except Exception as e:
            print('error reading the file!', e)
    
    def select_rows(self):

        self.input_table['Date'] = [str(self.input_table['Date'][i].date()) for i in range(len(self.input_table['Date']))]
        self.input_table['Time'] = [str(self.input_table['Time'][i]) for i in range(len(self.input_table['Time']))]

        self.input_table['Datetime'] = pd.to_datetime(self.input_table['Date'] + ' ' + self.input_table['Time'])
        self.input_table.set_index('Datetime', inplace=True)
        
        start_date = pd.to_datetime(self.start_date)
        end_date = pd.to_datetime(self.end_date)
        
        # Filter rows based on specified date range
        self.input_rows = self.input_table.loc[start_date:end_date]
        
        # Filter rows based on specified time range
        start_time = pd.to_datetime(self.start_time).time()
        end_time = pd.to_datetime(self.end_time).time()
        self.input_rows = self.input_rows.between_time(start_time, end_time).iloc[:, 2:]

        return self.input_rows
    
    def create_table1(self):
        if self.input_table is not None:
            
            for i, date in enumerate(self.target_dates):
                data = self.input_table[(self.input_table['Date'] == date) & (self.input_table['Time'] >= self.start_time) 
                                        & (self.input_table['Time'] <= self.end_time)]


                self.output_table1.loc[i] = [date.date(), data['OA TEMP'].mean(), data['OA RH'].mean()] + [np.round(data[self.input_column_names[j]].sum()/60) for j in range(4, 17)] + [np.round(data[self.input_column_names[20]].sum()/60)]

            self.output_table1.iloc[:, 1:] = self.output_table1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
            self.output_table1.loc[self.num_days] = ['合計'] +  [np.round(self.output_table1.loc[:, self.output_columns_names[j]].mean(), 2) for j in range(1, 3)] + [self.output_table1.loc[:, self.output_columns_names[j]].sum() for j in range(3, 17)]
            self.output_table1.iloc[self.num_days, 1:] = self.output_table1.iloc[self.num_days, 1:].apply(pd.to_numeric, errors='coerce')

        else:
            print("Failed to read the Excel file or date not found.")

    def create_table2(self):
        if self.output_table1 is not None:
            last_row = len(self.output_table1.index)
            column_groups = {'CH': [], 'PCHP': [], 'CDWP': [], 'CT': []}
            for column_name in self.output_columns_names:
                if 'PCHP' in column_name:
                    column_groups['PCHP'].append(column_name)
                elif 'CH' in column_name:
                    column_groups['CH'].append(column_name)
                elif 'CDWP' in column_name:
                    column_groups['CDWP'].append(column_name)
                elif 'CT' in column_name:
                    column_groups['CT'].append(column_name)

            system_efficiency_value = np.round(self.output_table1.at[last_row-1, '總累樍耗電量(kWh)'] / self.output_table1.at[last_row-1, '系統累積製冷能力(RT-H)'], 3)
            agreed_cold_energy_demand = 2203200
            improved_energy_consumption = system_efficiency_value*agreed_cold_energy_demand
            improved_energy_cost = improved_energy_consumption*2.6
            new_row = [self.output_table1.at[last_row-1, '系統累積製冷能力(RT-H)']] + [self.output_table1.loc[last_row-1, column_groups[group]].sum() for group in column_groups] + [system_efficiency_value, agreed_cold_energy_demand, improved_energy_consumption, improved_energy_cost]
            self.output_table2.loc[1] = new_row
        else:
            print("Failed to read the Excel file or date not found.")

    def create_table3(self):
        total_data = self.input_table.shape[0]
        hot_flat_horizontal_15_number_of_transactions = self.input_table.loc[:, self.input_column_names[-1]].sum()
        self.output_table3.loc[1] = [int(self.num_days), int(total_data), int(total_data), (total_data - total_data)/(1440*31), hot_flat_horizontal_15_number_of_transactions, hot_flat_horizontal_15_number_of_transactions/(1440*31), 1 - hot_flat_horizontal_15_number_of_transactions/(1440*31), self.output_table2.loc[1, "系統效率值(kW/RT)"]]

    def create_output_tables(self):
        self.extract_input_table()
        print('Creating table 1...')
        self.create_table1()
        print('Creating table 2...')
        self.create_table2()
        self.create_table3()
    
    def write_output_to_file(self):
        with pd.ExcelWriter(self.output_file_name) as writer:
            self.output_table1.to_excel(writer, sheet_name='Table 1')
            self.output_table2.to_excel(writer, sheet_name='Table 2')
            self.output_table3.to_excel(writer, sheet_name='Table 3')