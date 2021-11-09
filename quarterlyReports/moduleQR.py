#! python3

from pathlib import Path
from subprocess import Popen
from functools import reduce
from pprint import pprint
import os
import time

from openpyxl.utils import get_column_letter
import datefinder
import pandas as pd

from src.moduleInv import Inventory, df_to_xl

# pandas settings to display better output in shell
pd.set_option('display.max_rows', 30)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 500)


# Functions

def get_ns_path(dir):
    """
    Searches the Downloads directory for the .csv Inventory file
    downloaded directly from NetSuite, and chooses the most recent file created.
    """
    paths = []
    for folder_name, subfolders, filenames in os.walk(dir):
        for filename in filenames:
            if 'CSSeattleInventory' in filename:
                ti_c = os.path.getctime(dir / Path(filename))
                c_ti = time.ctime(ti_c)
                path = dir / Path(filename)
                paths.append((path, ti_c, c_ti))
    return sorted(paths, key=lambda x: x[1])[-1][0]

def get_ns_df(path, sku_map_path):
    """
    Returns a dataframe filled in with the correct details from sku_map and
    quantities from the recent inventory.csv file.
    """
    cols = ['Item', 'Customer SKU', 'Description', 'Category', 'Quantity']

    inv = Inventory(path)
    df = inv.report()
    df = df[cols[1:]]

    dg = pd.read_excel(sku_map_path)
    dg.rename(columns={'HB SKU' : 'Customer SKU'}, inplace=True)
    dg = dg[dg.columns[:2]]
    dg.dropna(inplace=True)
    dg.set_index(cols[0], inplace=True)

    df = df.combine_first(dg)
    df.dropna(subset=['Quantity'], inplace=True)

    return df[cols[1:]]

def get_current_snap(path_in, sku_map_path, path_out):
    """
    Saves and opens the detailed inventory snapshot.
    """
    df = get_ns_df(path_in, sku_map_path)
    df_to_xl(df, path_out, 'Snapshot')
    Popen(['open', path_out])

def og_snap_to_df(path):
    """
    Getting a dataframe from the 1-of, unique original inventory .xlsx file
    """
    df = pd.read_excel(path)
    df.rename(columns={'Quantity' : '2021-Q2 Qty'}, inplace=True)
    df.dropna(subset=['Item'], inplace=True)
    df.set_index('Item', inplace=True)
    return df[['Unit Cost', '2021-Q2 Qty']]

def get_snap_paths(dir):
    """
    Get the various paths of inventory snapshots in the snapshot directory,
    using a prescribed list of dates associated to the snapshot date.
    """
    paths = []
    for folder_name, subfolders, filenames in os.walk(dir):
        for filename in filenames:
            if ('Dean_Inventory' in filename and '07-06' not in filename):
                paths.append(dir / Path(filename))
    return paths

def snap_to_df(paths):
    """
    Turning the snapshot.xslx files into a list of dataframes.
    E.g., snap_to_df([file1.xlsx, file2.xlsx]) -> [df1, df2]
    """
    date_dict = {'01' : 'Q4',
                 '02' : 'Q1',
                 '03' : 'Q1',
                 '04' : 'Q1',
                 '05' : 'Q2',
                 '06' : 'Q2',
                 '07' : 'Q2',
                 '08' : 'Q3',
                 '09' : 'Q3',
                 '10' : 'Q3',
                 '11' : 'Q4',
                 '12' : 'Q4'}
    dfs = []
    for path in sorted(paths):
        date_raw = list(datefinder.find_dates(str(path)))[0]
        year = date_raw.strftime('%Y')
        month = date_raw.strftime('%m')
        quarter = date_dict[month]
        qty = year + '-' + quarter + ' Qty'

        df = pd.read_excel(path, index_col=0)
        if '2021-09-30' in str(path):
            df.reset_index(inplace=True)
            df = df[~df['Item'].astype(str).str.isdigit()]
            df.set_index('Item', inplace=True)
        dfs.append(df.rename(columns={'Quantity' : qty}))
    return dfs

def compare_snaps(dfs, path):
    """
    Given a list of dataframes, dfs, with dfs[0] being the original snapshot,
    compare_snaps returns a .xlsx file located at path.
    """
    df = reduce(lambda x, y: x.combine_first(y), dfs)

    cols = ['Customer SKU',
            'Description',
            'Category',
            'Unit Cost']
    qtys = [x for dg in dfs for x in list(dg.columns) if 'Qty' in x]
    cols = cols + qtys

    df = df[cols]
    df = df[df.index.notnull()]
    df[cols[5:]] = df[cols[5:]].fillna(0)
    df = df[df['Category'].notnull()]

    df_to_xl(df, path, 'Inventory Report', {2 : 0.8})
    Popen(['open', path])
    return df
