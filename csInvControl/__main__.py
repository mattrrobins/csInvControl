#! python3

from pathlib import Path
import sys
import os
import pprint
from inventoryClasses import Inventory

def get_dl_path(dir):
    for folder_name, subfolders, filenames in os.walk(dir):
        for filename in filenames:
            if 'CSSeattleInventory' in filename:
                return dir / Path(filename)

def search_check(searches):
    better_searches = []
    for search in searches:
        x = search.upper()
        if x == 'RECEIVING':
            x = x.title()
        elif x == 'LABREC':
            x = 'LabRec'
        better_searches.append(x)
    return better_searches



dl_dir = Path.home() / Path('Downloads')
up_dir = Path.home() / Path('Documents')
inv_report = up_dir / Path('inv_report.xlsx')
inv_list = up_dir / Path('temp_inv_list.xlsx')

floor_loc = ['Receiving',
             'GATE',
             'HOLD',
             'AISLE1',
             'AISLE2',
             'AISLE6',
             'PRODUCTION']
aisle_loc = ['1A', '1B', '2A', '2B', '3A', '3B', '4A', '5A', '5B', '6A']
label_loc = ['LL', 'LL2']
comp_loc = ['DIHALL', 'LabRec', 'FRIDGE',
                'M', 'N', 'O', 'P', 'U', 'V', 'W', 'X', 'Y', 'Z']
shipp_loc = ['SBOX', 'SMISC']



search = aisle_loc[2:3]

if __name__ == '__main__':
    try:
        inv = Inventory(get_dl_path(dl_dir))
        #inv.cycle_count(inv_list, search)
        inv.report_xl(inv_report)
    except ValueError:
        print('Forgot to download the csv file...')
