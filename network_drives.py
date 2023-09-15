
import os
from winreg import *


from ping3 import ping
import pandas as pd



from ldap3 import SUBTREE, Server, Connection


def getWinDrives():

    table = {"PC": [], "network drives": []}
    table1 = {"network drives": [], 'letter': [], 'PC': []}
    # table1 = {"network drives": []}
    # host = "BONDAR" # remote MACHINE or local
    AD_PASSWORD = 'Vfybgekzwbz3@!'
    AD_USER = 'optima-energy\ininsys'
    AD_SEARCH_TREE  = 'DC=optima-energy,DC=ru'
    # wmic /node:'BONDAR' /user:'optima-energy\ininsys' /PASSWORD:'Vfybgekzwbz3@!' service where (name='RemoteRegistry') get StartMode

    server = Server("192.168.1.4")
    conn = Connection(server, user=AD_USER, password=AD_PASSWORD)

    try:
        conn.bind()
        # все не отключенные PC, не DC, только windows
        conn.search(AD_SEARCH_TREE,
                    '(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',
                    SUBTREE, attributes=['cn'])
    except Exception as e:
        print(e)
        return False

    if not conn.entries:
        return False

    for entry in conn.entries:
        host = str(entry.cn)
        res0 = ping(host)
        if res0 is False:
            table["PC"].append(host)
            table["network drives"].append("ping Error")
            print(host, "ping Error")
            print('-----------------------------------')
            continue

        try:
            # получить тип запуска для службы windows удаленныый реестр
            servic_manager = os.popen(f"wmic /NODE:\"{host}\" /USER:\"{AD_USER}\" /PASSWORD:\"{AD_PASSWORD}\" service where (name=\'RemoteRegistry\') get StartMode",
                                          'r').read().split()
        except Exception as e:
            table["PC"].append(host)
            table["network drives"].append(e)
            print(host, e)
            print('-----------------------------------')
            continue

        if len(servic_manager) == 0:
            table["PC"].append(host)
            table["network drives"].append("Error. Отключимте firewall")
            print(host, "Error. Отключимте firewall")
            print('-----------------------------------')
            continue

        if 'Disabled' == servic_manager[1]:
            try:
                res = os.popen(f"wmic /NODE:\"{host}\" /USER:\"{AD_USER}\" /PASSWORD:\"{AD_PASSWORD}\" service where (name=\'RemoteRegistry\') CALL ChangeStartMode Automatic",
                            'r').read()
            except Exception as e:
                table["PC"].append(host)
                table["network drives"].append("Error CALL ChangeStartMode Automatic")
                print(host, "Error CALL ChangeStartMode Automatic")
                print('-----------------------------------')
                continue

            if 'ReturnValue = 0' in res:
                res1 = os.popen(f"wmic /NODE:\"{host}\" /USER:\"{AD_USER}\" /PASSWORD:\"{AD_PASSWORD}\" service where (name=\'RemoteRegistry\') CALL StartService",
                                'r').read()
                if 'ReturnValue = 0' not in res1:
                    table["PC"].append(host)
                    table["network drives"].append(" Error CALL StartService")
                    print(host, " Error CALL StartService")
                    print('-----------------------------------')
                    continue
            else:
                table["PC"].append(host)
                table["network drives"].append(res)
                print(host, res)
                print('-----------------------------------')
                continue

        try:
            key = ConnectRegistry(host, HKEY_USERS)
        except FileNotFoundError as  e:
            table["PC"].append(host)
            table["network drives"].append(e)
            print(host, e)
            print('-----------------------------------')
            continue

        hkeys = OpenKey(key, '')
        num = QueryInfoKey(hkeys)[0]
        i = 0
        print(host)
        listtmp = []
        while i < num:
            res = EnumKey(hkeys, i)
            if res == ".DEFAULT":
                i += 1
                continue
            try:
                hkeys1 = OpenKey(key, f"{res}\\Network")
                num1 = QueryInfoKey(hkeys1)[0]
                if num1:
                    i1 = 0
                    while i1 < num1:
                        res1 = EnumKey(hkeys1, i1)
                        hkeys2 = OpenKey(key, f"{res}\\Network\\{res1}")
                        i2 = 0
                        while True:
                            try:
                                name, value, value_type = EnumValue(hkeys2, i2)
                                if name == 'RemotePath':
                                    r = ''.join([res1,':', value, ','])
                                    listtmp.append(r)
                                    try:
                                        iii = table1["network drives"].index(value)
                                        if res1 not in table1['letter'][iii]:
                                            table1['letter'][iii].append(res1)
                                        if host not in table1['PC'][iii]:
                                            table1['PC'][iii].append(host)
                                    except ValueError as e:
                                        table1["network drives"].append(value)
                                        table1['letter'].append([res1])
                                        table1['PC'].append([host])

                            except Exception as e:
                                break
                            i2 += 1
                        i1 += 1

            except Exception as e:
                pass
            i += 1
        if not listtmp: listtmp.append('Free')
        listtmp = " ".join(listtmp)
        print(listtmp)
        print(table1)
        print('-----------------------------------')
        table["PC"].append(host)
        table["network drives"].append(listtmp)

        res2 = os.popen(f"wmic /NODE:\"{host}\" /USER:\"{AD_USER}\" /PASSWORD:\"{AD_PASSWORD}\" service where (name=\'RemoteRegistry\') CALL StopService",
                                'r').read()
        if "Disabled" == servic_manager[1]:
            res3 = os.popen(f"wmic /NODE:\"{host}\" /USER:\"{AD_USER}\" /PASSWORD:\"{AD_PASSWORD}\" service where (name=\'RemoteRegistry\') CALL ChangeStartMode {servic_manager[1]}",
                            'r').read()
    return table, table1




if __name__ == '__main__':
    table, table1 = getWinDrives()
    if table:
        df = pd.DataFrame(table)
        df1 = pd.DataFrame(table1)
        df.to_excel("table.xlsx")
        df1.to_excel("network_drives.xlsx")
    else:
        print("Ошибка сервера LADP")