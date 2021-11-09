#! python3

from pathlib import Path
from datetime import date
import sys

from moduleQR import *

dl_dir = Path.home() / Path('Downloads')
proj_dir = Path(__file__).resolve().parents[1]


sku_map_path = proj_dir / Path('masterLists', 'herbivoreMasterList',
                                'unofficial_sku_map.xlsx')

snap_dir = proj_dir / Path('quarterlyReports', 'snapshots')
today = str(date.today())
cur_snap = 'Dean_Inventory_%s.xlsx' % today
cur_snap_path = snap_dir / Path(cur_snap)
inv_report = 'CS_Inventory_Quarterly_Report_%s.xlsx' % today
inv_report_path = snap_dir / Path(inv_report)



def main():
    try:
        opt = sys.argv[1].lower()
        if opt == 'snap':
            get_current_snap(get_ns_path(dl_dir), sku_map_path, cur_snap_path)
        elif opt == 'compare':
            df_0 = og_snap_to_df(Path(snap_dir, 'Dean_Inventory_2021-07-06.xlsx'))
            dfs = snap_to_df(get_snap_paths(snap_dir))
            df = compare_snaps([df_0] + dfs, inv_report_path)
        else:
            print('This program requires an argument of either "snap" or "compare".')
    except IndexError:
        print('This program requires an argument of either "snap" or "compare".')





if __name__ == '__main__':
    main()
