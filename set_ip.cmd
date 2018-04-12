REM Set IP addresses
echo off
cls
set /p ip=Enter IP address:
echo Setting IP address, Subnet and Gateway...
netsh interface ip set address name="Local Area Connection" static %ip% 255.255.255.128 10.65.58.1
echo .
set /p dns1=Enter DNS 1 address:
set /p dns2=Enter DNS 2 address:
echo Setting DNS Settings
netsh interface ip add dns name="Local Area Connection" addr=%dns1%
netsh interface ip add dns name="Local Area Connection" addr=%dns2% index=2


