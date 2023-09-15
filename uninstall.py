import os

import pandas as pd
from ping3 import ping

P = 'Vfybgekzwbz3@!'
U = 'optima-energy\ininsys'
AD_SEARCH_TREE  = 'DC=optima-energy,DC=ru'

def uninstall():
    df = pd.read_excel("ESET.xlsx")
    res = df.loc[(df["IP"].str.startswith("192.168.1.")) & (df["PC"] != "UVT")]
    # wmic /node:"SHVETCOV.OPTIMA-ENERGY.RU" product where name="ESET Endpoint Security" call uninstall
    res2 = res[['PC', 'installed progs']]
    for index, row in res2.iterrows():
        res3 = ping(row['PC'])
        if res3 is False:
            print(f'Нет пинга: {row["PC"]}')
            print('==========================')
            continue
        try:
            host = str(row['PC'])
            prod = str(row['installed progs']).strip()

            cmd = f"wmic /NODE:\"{host}\" /USER:\"{U}\" /PASSWORD:\"{P}\" product where \"Name like '{prod}'\" call uninstall"
            print(f"========={cmd}===============")
            res1 = os.popen(cmd, 'r').read()
            print(res1)
        except Exception as e:
            print(f'{row["PC"]} ERROR FOR {e}')
        print('-------------------------------------------')
        # print(row['PC'], row['installed progs'])


if __name__ == '__main__':
    uninstall()
