@echo off
REM - Batch file copies over the AutoDiscover file 
REM - to fix the Outlook 2016 and Office 365 connection
REM - Phillip Trimble, COT, Last updated: 2/26/18

cls
echo Please Type the Users login name i.e. john.doe:
set /p user=

cls
if not exist c:\autodiscovery mkdir c:\autodiscovery
copy "\\eas.ds.ky.gov\dfs\DS_Share\Admins Only\Utilities\O365_Autodiscover\Autodiscover.xml" c:\autodiscovery\
copy "\\eas.ds.ky.gov\dfs\DS_Share\Admins Only\Utilities\O365_Autodiscover\o365_autodiscover_reg_fix.reg" c:\users\%user%\desktop\

cls
echo Please have user login and run the o365_autodiscover_reg_fix.reg from their desktop!
echo .
echo .
pause


