import os
from winreg import *

from ping3 import ping
import pandas as pd

from ldap3 import SUBTREE, Server, Connection


def getWinDrives():
    table0 = {"PC": [], "network drives": []}
    table1 = {"network drives": [], 'letter': [], 'PC': []}
    # table1 = {"network drives": []}
    # host = "BONDAR" # remote MACHINE or local
    AD_PASSWORD = 'Vfybgekzwbz3@!'
    AD_USER = 'domain2\ininsys'
    AD_SEARCH_TREE = 'DC=domain2,DC=local'
    # wmic /node:'' /user:'' /PASSWORD:'!' service where (name='RemoteRegistry') get StartMode

    server = Server("192.168.1.7")
    conn = Connection(server, user=AD_USER, password=AD_PASSWORD)

    try:
        conn.bind()
        # все не отключенные PC, не DC, только windows
        conn.search(
            AD_SEARCH_TREE,
            '(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',
            SUBTREE,
            attributes=['cn'])
    except Exception as e:
        print(e)
        exit(1)

    if not conn.entries:
        print(f"сервер: {server} не ответил")
        exit(1)

    for entry in conn.entries:
        host = str(entry.cn)
        res0 = ping(host)
        if res0 is False:
            table0["PC"].append(host)
            table0["network drives"].append("ping Error")
            print(host, "ping Error")
            print('-----------------------------------')
            continue
        else:
            print(host)
            try:
                path = f"\\\\{host}\c$\Windows\System32\drivers\etc\hosts"
                print(path)
                res11 = os.popen(f"echo 127.0.0.1 ase.autodesk.com >> {path}").read().split()
                print(res11)
            except Exception as e:
                print(e)
            print('-----------------------------------')


if __name__ == '__main__':
    getWinDrives()

