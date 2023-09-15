@echo off
chcp 1251

set str=%username%.%COMPUTERNAME%
WMIC LOGICALDISK where DriveType=4 get DeviceID, ProviderName > %str%.txt
pause


