import os.path
import sys
import matplotlib.pyplot as pplt
import numpy as np
import pandas as pd
import openpyxl
from scipy import stats

#Define source
location = 'C:\\Users\\970jwillems\\Sonova\\Sonova Marketing GmbH - General\\05_DA\\03_Tickets\\Janine\\02_DE_CuCa\\DASD-456 Q-Screen Deep-Dive Bonus Report\\'
file =  'dasd-456_report.xlsx'
filepath = os.path.join(location, file)
df1 = pd.read_excel(filepath, sheet_name = "Raw Data")

#Create calculated columns
df1['qscreen_ratio'] =  df1['num_cust_qscreen'] / df1['num_cust_sched_app']
df1['calls_per_hour'] = df1['num_outbound_calls'] / df1['worked_hours']
df1['tickets_per_hour'] = df1['num_closed_lead_tickets'] / df1['worked_hours']
df1['aht'] = df1['call_handling_time_outbound_reached'] / df1['num_outbound_calls_reached']
df1['occupancy'] = (df1['call_handling_time_outbound'] + df1['call_handling_time_inbound']) / (df1['call_handling_time_outbound'] + df1['call_handling_time_inbound'] + df1['ready_time'])
df1['reached_rate'] = df1['num_outbound_calls_reached'] / df1['num_outbound_calls']

agents = ['Niklas Wefeld', 'Kathrin Beilfuß', 'Nicole Starke', 'Johanna Schäfer']
metrics = ['qscreen_ratio', 'aht', 'calls_per_hour', 'tickets_per_hour']                                  #for T-Test
ys = ['occupancy', 'ready_time', 'aht', 'calls_per_hour', 'tickets_per_hour', 'reached_rate']   #for Correlations

#T-Test - Performance Q-Screen Ratio Comparison Last 10 Weeks
print('Last 10 Weeks Comparison', '\n')
df_tp1 = df1.where(df1['last_10_weeks_flag'] == 0)
df_tp2 = df1.where(df1['last_10_weeks_flag'] == 1)

for agent in agents:
    for metric in metrics:
        df_tp1_ag = df_tp1.where(df_tp1['agent_name']==agent)
        df_tp2_ag = df_tp2.where(df_tp2['agent_name']==agent)
        print('Agent:', agent, 'Metric', metric)
        ##print('Mean 1:', ((df_tp1_ag['num_cust_qscreen'].sum() / df_tp1_ag['num_cust_sched_app'].sum()).mean()),'Mean 2:', ((df_tp2_ag['num_cust_qscreen'].sum() / df_tp2_ag['num_cust_sched_app'].sum()).mean()))
        print('Mean 1:', df_tp1_ag[metric].mean(), 'Mean 2:', df_tp2_ag[metric].mean())
        print('Shapiro Test 1:', stats.shapiro(df_tp1_ag[metric].dropna()), 'Shapiro Test 2:', stats.shapiro(df_tp2_ag[metric].dropna()))
        print('Independent T-Test Result:', stats.ttest_ind(df_tp1_ag[metric].dropna(), df_tp2_ag[metric].dropna()))
        print('Welchs T-Test Result:', stats.ttest_ind(df_tp1_ag[metric].dropna(), df_tp2_ag[metric].dropna(), equal_var=False), '\n')
        df_tp1_ag = df_tp1
        df_tp2_ag = df_tp2

#T-Test - Performance Q-Screen Ratio Comparison Last 16 Weeks
print('Last 16 Weeks Comparison', '\n')
df_tp1 = df1.where(df1['last_16_weeks_flag'] == 0)
df_tp2 = df1.where(df1['last_16_weeks_flag'] == 1)

for agent in agents:
    for metric in metrics:
        df_tp1_ag = df_tp1.where(df_tp1['agent_name']==agent)
        df_tp2_ag = df_tp2.where(df_tp2['agent_name']==agent)
        print('Agent:', agent, 'Metric', metric)
        print('Mean 1:', df_tp1_ag[metric].mean(), 'Mean 2:', df_tp2_ag[metric].mean())
        print('Shapiro Test 1:', stats.shapiro(df_tp1_ag[metric].dropna()), 'Shapiro Test 2:', stats.shapiro(df_tp2_ag[metric].dropna()))
        print('Independent T-Test Result:', stats.ttest_ind(df_tp1_ag[metric].dropna(), df_tp2_ag[metric].dropna()))
        print('Welchs T-Test Result:', stats.ttest_ind(df_tp1_ag[metric].dropna(), df_tp2_ag[metric].dropna(), equal_var=False), '\n')
        df_tp1_ag = df_tp1
        df_tp2_ag = df_tp2

#Get correlatios
x = 'qscreen_ratio'

class Correlation:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def rscore(self):
        rscore = df[x].corr(df[y], method='pearson')
        print('Correlation between', f'{x}', 'and', f'{y}:', rscore, '\n')

for agent in agents:
    for y in ys:
        df = df1.where(df1['agent_name']==agent)
        print('Agent:', agent)
        Correlation(x,y).rscore()
        df = df



















    
