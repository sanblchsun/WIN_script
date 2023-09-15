import asyncio
import os
import subprocess
import sys
import logging
import socket
import time

import pandas as pd
from ldap3 import SUBTREE, Server, Connection
import loader

logging.basicConfig()


def connect_ldap():
    ad_password = loader.ad_password
    ad_user = loader.ad_user
    ad_search_tree = loader.ad_search_tree
    server0 = loader.server
    server = Server(server0)
    conn = Connection(server, user=ad_user, password=ad_password)

    try:
        conn.bind()
        # все не отключенные PC, не DC, только windows
        conn.search(ad_search_tree,
                    '(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',
                    SUBTREE, attributes=['cn'])
    except Exception as e:
        print(f"Ошибка LDAP!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!{e}")
        sys.exit(0)
    if not conn.entries:
        print("соединение имеет False")
        sys.exit(0)
    return conn, ad_user, ad_password


async def run_wmic(cmd):
    await asyncio.sleep(0.5)
    func_return = os.popen(cmd).read()
    return func_return


async def get_ip(host):
    await asyncio.sleep(0.5)
    try:
        return socket.gethostbyname(host)
    except socket.gaierror:
        return False


async def execute_task(host, ad_user, ad_password, table):
    res_list = []
    ip_host = await get_ip(host)
    if not ip_host:
        res_list.append('Не пигуется')
    else:
        res_list.append(ip_host)
        cmd = f"wmic /NODE:\"{host}\" /USER:\"{ad_user}\" /PASSWORD:\"{ad_password}\" product get name"
        res1 = await run_wmic(cmd)
        if res1 == "\n\n\n\n":
            res_list.append("wmic False")
        else:
            list_install = list(map(str.strip, res1.replace('\n\n', ',').split(',')))
            res_list.append("ESET не найден")
            for i in list_install:
                if "ESET" in i and "Security" in i:
                    res_list.clear()
                    res_list.append(i)
                    table["PC"].append(host)
                    table["IP"].append(ip_host)
                    table["installed progs"].append(i)

    print(f'{res_list}\n============================{host}====================')


async def main():
    table = {"PC": [], "IP": [], "installed progs": []}
    conn, ad_user, ad_password = connect_ldap()
    lst = []
    for entry in conn.entries:
        host = str(entry.cn)
        task = asyncio.create_task(execute_task(host, ad_user, ad_password, table))
        lst.append(task)
    await asyncio.gather(*lst)
    if table:
        df = pd.DataFrame(table)
        df.to_excel("ESET.xlsx")
    else:
        print("Ошибка таблицы")


if __name__ == '__main__':
    t0 = time.time()
    asyncio.run(main())
    print(int(time.time() - t0) / 60)
