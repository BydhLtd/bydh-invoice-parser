#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import numpy as np
import re


# In[2]:


CURRENT_DIR = os.getcwd()
FILES = []

for file in os.listdir(CURRENT_DIR):
    if re.search('^sample(.*?).xlsx',file):
        FILES.append(file)


# In[3]:


def load_preprocessing(input_dir):
    raw_data_df = pd.read_excel(input_dir)
    select_columns = ['Shear rate','Torque', 'Stress']
    preprocessed_data_df = raw_data_df.copy()
    preprocessed_data_df = preprocessed_data_df.drop(index=0, axis=0)
    return preprocessed_data_df


# In[4]:


def process_data(preprocessed_data_df):
    output_df_dict = {'Shear rate 1/s':[], 'Torque µN.m':[], 'Torque N.m':[], 'Stress MPa':[], 'Stress Pa':[]}
    
    for i in range(preprocessed_data_df.shape[0]):
        row = preprocessed_data_df.iloc[i]
        share_rate, torque_unm, stress_mpa = row['Shear rate'],row['Torque'], row['Stress']
        torque_nm, stress_pa =  torque_unm * (10**-6), stress_mpa * (10**6)
        output_df_dict['Shear rate 1/s'].append(np.around(share_rate,decimals=1))
        output_df_dict['Torque µN.m'].append(torque_unm)
        output_df_dict['Torque N.m'].append(torque_nm)
        output_df_dict['Stress MPa'].append(stress_mpa)
        output_df_dict['Stress Pa'].append(stress_pa)
    
    output_df = pd.DataFrame(output_df_dict)
    
    return output_df


# In[5]:


for file in FILES:
    file_dir = os.path.join(CURRENT_DIR, file)
    output_dir = os.path.join(CURRENT_DIR, 'parsed ' + file)

    preprocessed_data_df = load_preprocessing(file_dir)
    output_df = process_data(preprocessed_data_df)
    output_df.to_excel(output_dir,index=False)

