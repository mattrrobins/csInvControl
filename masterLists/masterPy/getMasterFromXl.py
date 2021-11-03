#! python3

from pathlib import Path
from subprocess import Popen
import pprint

import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter

from csInvControl.inventoryClasses import df_to_xl

proj_dir = Path(__file__).resolve().parents[1]
ml_dir = proj_dir / Path('herbivoreMasterList')
hml_xl = Path(ml_dir, 'herbivore_master_list.xlsx')
sku_xl = Path(ml_dir, 'unofficial_sku_map.xlsx')

cat_map = {'Finished Goods' : 'Finished Goods',
           'WIP' : 'WIP',
           'Bulk' : 'Bulk',
           'Containers' : 'Containters',
           'Closures' : 'Closures',
           'Sealing Disks' : 'Sealing Disks',
           'Flutes' : 'Flutes',
           'Cartons' : 'Cartons',
           'Bags' : 'Bags',
           'Kit Components' : 'Kit Components',
           'Machine Labels' : 'Labels',
           'UPCs' : 'Labels',
           'INCIs' : 'Labels',
           'Stickers' : 'Labels',
           'Essential Oils' : 'Raw Materials',
           'Dry Goods' : 'Raw Materials',
           'Carriers' : 'Raw Materials',
           'Shipping Boxes' : 'Shipping Boxes',
           'Case Packs' : 'Case Packs',
           'Partitions' : 'Partitions'}

def xl_to_df(path):
    """
    Pulls a .xlsx file and converts to the information into a more condensed
    .xlsx file for distribution relating the two companies SKUs to each other.
    """
    xl = pd.ExcelFile(path)
    sheets = xl.sheet_names
    cols = ['HB SKU',
            'SC Part Number',
            'Description',
            'Category']
    df = pd.concat([pd.read_excel(xl, sheet_name=s)
                    .assign(Category=cat_map[s]) for s in sheets],
                    ignore_index=True)
    df.set_index('Item', inplace=True)
    df = df[cols]

    return df


if __name__ == '__main__':
    df = xl_to_df(hml_xl)
    df_to_xl(df, sku_xl, 'HB SKU to CS Item Map', {1 : 1.5})
    Popen(['open', sku_xl])
