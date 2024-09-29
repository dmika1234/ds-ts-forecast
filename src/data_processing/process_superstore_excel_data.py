import sys
import os
# Adding local import paths
src_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '../'))
os.chdir(src_dir)
sys.path.append(src_dir)
# Local imports
from config import *
# Standard imports
import pandas as pd


# Load excel data
excel_data_file_name = 'Superstore.xls'
orders_df = pd.read_excel(os.path.join(DATA_DIR, excel_data_file_name), sheet_name='Orders')
returns_df = pd.read_excel(os.path.join(DATA_DIR, excel_data_file_name), sheet_name='Returns')
people_df = pd.read_excel(os.path.join(DATA_DIR, excel_data_file_name), sheet_name='People')
# Merge data
final_store_df = pd.merge(orders_df, returns_df, how='left', on='Order ID', suffixes=(' Order ', ' Return'))
final_store_df = pd.merge(final_store_df, people_df, how='left', left_on='Customer Name', right_on='Person', suffixes=(' Order ', ' Customer'))
# Save data
final_store_df.to_csv(os.path.join(DATA_DIR, 'superstore_data.csv'), index=False)