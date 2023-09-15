# pip install pandas openpyxl xlsxwriter xlrd ldap3
import os

from PyQt6 import QtCore
# from ldap3 import Server, Connection, SIMPLE, SYNC, ASYNC, SUBTREE, ALL
from ldap3 import Server, Connection, SUBTREE
import pandas as pd


# Object, which will be moved to another thread
class BrowserHandler(QtCore.QObject):
    running = False
    newTextAndColor = QtCore.pyqtSignal(str)
    progressBarValue = QtCore.pyqtSignal(int)
    def __init__(self, ip_addres, ad_user, ad_passord, ad_search_tree, file_name):
        super().__init__()
        self.ip_addres = ip_addres
        self.AD_PASSWORD = ad_passord
        self.AD_USER = ad_user
        self.AD_SEARCH_TREE = ad_search_tree
        self.file_name = file_name


    # method which will execute algorithm in another thread
    def run(self):

        self.server = Server(self.ip_addres)
        self.conn = Connection(self.server, user=self.AD_USER, password=self.AD_PASSWORD)

        try:
            self.conn.bind()
            # все не отключенные ПК, не DC, только windows
            self.conn.search(self.AD_SEARCH_TREE,
                             '(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',
                             SUBTREE, attributes=['cn'])
            self.table = {"pc": [], "speed Мбит/сек": []}

            cnt = len(self.conn.entries)
            iii = 0
            for entry in self.conn.entries:
                iii += 1
                i = iii*100//cnt
                self.progressBarValue.emit(i)
                pc = str(entry.cn)
                self.table['pc'].append(pc)
                self.speed = os.popen(
                    f"wmic /NODE:\"{pc}\" /USER:\"{self.AD_USER}\" /PASSWORD:\"{self.AD_PASSWORD}\" NIC where \"NetEnabled=\'true\'\" get \"Speed\"",
                    'r').read()
                self.speed = self.speed.replace("Speed", "").strip()
                if self.speed:
                    self.speed = str(int(self.speed)/1000000)
                    self.newTextAndColor.emit(f"{iii}) ПК: {pc}  speed: {self.speed} Мбит/сек")
                else:
                    self.speed = "Сервер RPC недоступен."
                    self.newTextAndColor.emit(f"{iii}) ПК: {pc}  speed: {self.speed}")

                self.table['speed Мбит/сек'].append(self.speed)

            self.newTextAndColor.emit(f"Выполнено 100 %, это {iii} ПК")
            df = pd.DataFrame(self.table)
            df.to_excel(self.file_name)
        except Exception as e:
            print(f"Ошибка соединения с сервером {self.ip_addres} (ошибка: {e})")


        # for pc in self.table['pc']:
        #     self.newTextAndColor.emit(pc)
        #     speed = os.popen(f"wmic /NODE:\"{pc}\" NIC where \"NetEnabled=\'true\'\" get \"Speed\"", 'r').read()
        #     speed = speed.replace("Speed", "").strip()
        #     self.table['speed'].append(speed)

        # while True:
        #
        #     # send signal with new text and color from aonther thread
        #     self.newTextAndColor.emit(
        #         '{} - thread 2 variant 1.\n'.format(str(time.strftime("%Y-%m-%d-%H.%M.%S", time.localtime()))),
        #         QColor(0, 0, 255)
        #     )
        #     QtCore.QThread.msleep(1000)
        #
        #     # send signal with new text and color from aonther thread
        #     self.newTextAndColor.emit(
        #         '{} - thread 2 variant 2.\n'.format(str(time.strftime("%Y-%m-%d-%H.%M.%S", time.localtime()))),
        #         QColor(255, 0, 0)
        #     )
        #     QtCore.QThread.msleep(1000)

# class Request_speed(QtCore.QObject):
#     running = False
#
#     def run(self):
#         for i in range(1, 100):
#             print(i)
# QtCore.QThread.msleep(1000)

# for pc in self.table['pc']:
# speed = os.popen(f"wmic /NODE:\"{pc}\" NIC where \"NetEnabled=\'true\'\" get \"Speed\"", 'r').read()
# speed = '1111'
# self.mainwindow.condition.setText(pc)
# speed = speed.replace("Speed", "").strip()
# self.table['speed'].append(speed)


# class Lst_speed(QtCore.QObject):
#     def __init__(self, ip_addres, domain, admin, passwd, filename):
#         super().__init__()
#         self.AD_SERVER = ip_addres
#         self.AD_SEARCH_TREE = domain
#         self.AD_USER = admin
#         self.AD_PASSWORD = passwd
#         self.FILE_NAME = filename
#         self.server = Server(self.AD_SERVER)
#         self.conn = Connection(self.server, user=self.AD_USER, password=self.AD_PASSWORD)
#
#
#     def setconnect(self):
#         try:
#             self.conn.bind()
#             # все не отключенные ПК, не DC, только windows
#             self.conn.search(self.AD_SEARCH_TREE,'(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',SUBTREE,attributes =['cn'])
#         except Exception as e:
#             return f"Ошибка соединения с сервером {self.AD_SERVER} (ошибка: {e})"
#
#     def get_adapter_speed(self):
#         table = {"pc": [], "speed": []}
#         for entry in self.conn.entries:
#             pc = str(entry.cn)
#             table['pc'].append(pc)
#         return table
#         # df = pd.DataFrame(table)
#         # df.to_excel('who_is_who.xlsx')
#
#     def run(self):
#         self.setconnect()
#         res = self.get_adapter_speed()
#         return res


# conn1 = setconnect(AD_SERVER="192.168.1.4", AD_USER=AD_USER, AD_PASSWORD=AD_PASSWORD)
# conn2 = setconnect(AD_SERVER="192.168.200.12", AD_USER=AD_USER, AD_PASSWORD=AD_PASSWORD)


# conn1.search(AD_SEARCH_TREE, '(&(objectCategory=computer))', SUBTREE, attributes =['cn'])
# коньроллеры домена
# conn.search(AD_SEARCH_TREE,'(&(objectCategory=computer)(primaryGroupID=516))')
# Операционка windows
# conn.search(AD_SEARCH_TREE,'(&(objectCategory=computer)(operatingSystem=*windows*))')
#  без DC
# conn.search(AD_SEARCH_TREE,'(&(objectCategory=computer)(primaryGroupID=515))')
#  все не отключенные ПК, не DC, только windows
# conn.search(AD_SEARCH_TREE,'(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=8192))(operatingSystem=*windows*))',SUBTREE,attributes =['cn'])
# conn.search(AD_SEARCH_TREE,'(&(objectCategory=Person)(sAMAccountName=buh))', SUBTREE,
#     attributes =['cn','proxyAddresses','department','sAMAccountName', 'displayName', 'telephoneNumber', 'ipPhone', 'streetAddress',
#     'title','manager','objectGUID','company','lastLogon']
#     )


# print(conn.entries)
#
# table = {"pc": [], "speed": []}
#
# for entry in conn.entries:
#    pc = str(entry.cn)
#    table['pc'].append(pc)
#    print(pc)
#    speed = os.popen(f"wmic /NODE:\"{pc}\" NIC where \"NetEnabled=\'true\'\" get \"Speed\"").read()
#    speed = speed.replace("Speed", "").strip()
#    print(speed)
#    table['speed'].append(speed)
#    print('------------------------------')
#
#
# df = pd.DataFrame(table)
# df.to_excel('who_is_who.xlsx')


# table1 = {"NAME": [], "GUID": []}
#
# for entry in conn1.entries:
#    name = entry.sAMAccountName
#    table1['NAME'].append(name)
#    guid = entry.objectGUID
#    table1['GUID'].append(guid)
#
# table2 = {"NAME": [], "GUID": []}
#
# for entry in conn2.entries:
#    name = entry.sAMAccountName
#    table2['NAME'].append(name)
#    guid = entry.objectGUID
#    table2['GUID'].append(guid)
#
#
# for iter in table1['NAME']:
#     pos0 = table1['NAME'].index(iter)
#     pos = -1
#     try:
#         pos = table2["NAME"].index(iter)
#     except Exception as e:
#         print(f'нериплицировался: {iter} ошибка: {e}')
#     if pos>-1:
#         if table1['GUID'][pos0]!=table2['GUID'][pos]:
#             name = table1['NAME'][pos]
#             print(f'не совпадают GUID для: {name}')

#
# df = pd.DataFrame(table1)
# df.to_excel('DC_trable.xlsx')
