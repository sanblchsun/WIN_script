@echo off
chcp 1251

set str=%username%.%COMPUTERNAME%
WMIC product get name > %str%.product.txt
WMIC LOGICALDISK where DriveType=4 get DeviceID, ProviderName > %str%.DISK.txt
pause


