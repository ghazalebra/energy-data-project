import pandas as pd
import numpy as np



class Data:

    def __init__(self, file=None, start_date='2023-08-01', end_date='2023-08-02', output_file_name="output.xlsx"):
        self.file = file
        self.start_date = start_date
        self.end_date = end_date
        self.output_file_name = output_file_name
    
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

    # start_date = pd.to_datetime(input('Enter the start date please: ') or '2023-8-1')
    # end_date = pd.to_datetime(input('Enter the end date please: ') or '2023-8-31')

    output_table1 = pd.DataFrame(columns=output_columns_names)
    output_table2 = pd.DataFrame(columns=output_column_names2)
    output_table3 = pd.DataFrame(columns=output_columns_names3)

    def extract_input_table(self):
        try:
            df = pd.read_excel(self.file, skiprows=[0], dtype={'OA TEMP': float})
            self.input_table = df.iloc[:, :self.num_cols]
            self.input_table['Date'] = pd.to_datetime(self.input_table['Date']).dt.date
        except Exception as e:
            print('error reading the file!', e)
            # return("An error occurred while reading the data:", e)

    def find_date_range(self):
        self.target_dates = [date.date() for date in pd.date_range(self.start_date, self.end_date, freq='d').tolist()]

    def select_one_day(self, target_date):
        """
        Selects data corresponding to 1 day
        """
        try:

            dates = self.input_table["Date"]
            
            # Select rows with the target data
            target_rows = self.input_table[dates == target_date]
            
            return target_rows
        
        except Exception as e:
            print("An error occurred:", e)
            return None
    
    def select_days(self):
        self.input_rows = {target_date: self.select_one_day(target_date) for target_date in self.target_dates}
        
    def create_table1(self):
        if self.input_rows is not None:
            dates = self.input_rows.keys()
            for i, date in enumerate(dates):

                data = self.input_rows[date]
                self.output_table1.loc[i] = [date, data['OA TEMP'].mean(), data['OA RH'].mean()] + [np.round(data[self.input_column_names[j]].sum()/60) for j in range(4, 17)] + [np.round(data[self.input_column_names[20]].sum()/60)]

            self.output_table1.iloc[:, 1:] = self.output_table1.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
            # print(output_table1)
            self.output_table1.loc[len(dates)] = ['合計'] +  [np.round(self.output_table1.loc[:, self.output_columns_names[j]].mean(), 2) for j in range(1, 3)] + [self.output_table1.loc[:, self.output_columns_names[j]].sum() for j in range(3, 17)]
            self.output_table1.iloc[len(dates), 1:] = self.output_table1.iloc[len(dates), 1:].apply(pd.to_numeric, errors='coerce')

        else:
            print("Failed to read the Excel file or date not found.")

    def create_table2(self):
        if self.output_table1 is not None:
            last_row = len(self.output_table1.index)
            # columns containing CH
            column_groups = {'CH': [], 'PCHP': [], 'CDWP': [], 'CT': []}

            for column_name in self.output_columns_names:
                if 'CH' in column_name:
                    column_groups['PCHP'].append(column_name)
                elif 'PCHP' in column_name:
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
        self.output_table3.loc[1] = [len(self.target_dates), total_data, total_data, (total_data - total_data)/(1440*31), hot_flat_horizontal_15_number_of_transactions, hot_flat_horizontal_15_number_of_transactions/(1440*31), 1 - hot_flat_horizontal_15_number_of_transactions/(1440*31), self.output_table2.loc[1, "系統 效率值"]]

    def create_output_tables(self):
        self.extract_input_table()
        self.find_date_range()
        print('Selecting the rows...')
        self.select_days()
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