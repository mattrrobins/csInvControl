#! python3

from pathlib import Path
from subprocess import Popen
import pprint

import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter


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

def df_to_xl(df, path, sheet_name, w={}):
    col_size = []
    col_size.append(max(df.index.astype(str).map(len)))
    for col in df.columns:
        m = max(max(df[col].astype(str).map(len)), len(str(col)))
        col_size.append(m)
    for k in w:
        col_size[k] *= w[k]

    with pd.ExcelWriter(path, mode='w', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name)
        ws = writer.sheets[sheet_name]
        for i in range(1, ws.max_column + 1):
            let = get_column_letter(i)
            ws.column_dimensions[let].width = col_size[i - 1]




df = xl_to_df(hml_xl)
df_to_xl(df, sku_xl, 'HB SKU to CS Item Map', {1 : 1.5})
Popen(['open', sku_xl])
pprint.pprint(df)
