<#
.SYNOPSIS
Installs the RICPAY software and all the related requirements.
.DESCRIPTION
This script install the RICPAY software.  This includes ScandAll, scanner drivers
PSFTP, Header Sheet and Bar Code Font.
.PARAMETER Name
User is prompted for all parameters in real time.
.PARAMETER Extension
User is prompted for all parameters in real time.
.INPUTS
User will supply the County Name "Johnson" and County Code "058".
.OUTPUTS
No Output is returned other than write-host.
.EXAMPLE
Copy the following 2 files over to the destination computer's desktop
_ricpay_install_full_deploy-v4.2.cmd and _ricpay_install_full_deploy-v4.2.ps1
.EXAMPLE
To execute the script - right click on _ricpay_install_full_deploy-v4.2.cmd and choose
Run As Administrator.
Last updated: 1/10/2024 - Author: Phillip Trimble, COT - Version 4.23
Misc: Auto launch ScandAll with Batch - "c:\program files (x86)\fiScanner\Scandall Pro\ScandallPro.exe" /BATCH:RICPAYJOHNSON041514 /Exit
#>

$global:rc = ""
$global:bt = ""
$ErrorActionPreference = "SilentlyContinue"

function global:inst-ricpay-local
    
  {
    # Copy Ricpay county specific files to local computer...
    # _____________________________________________________________________________________
    write-host ""
    write-host "Copying County specific files to $env:COMPUTERNAME" -ForegroundColor Magenta
    New-Item "c:\Ricpay" -type directory -force
    copy-item "\\$inst\RICPAY - CHFS\$county\code39.ttf" "c:\ricpay\" -force 
    copy-item "\\$inst\RICPAY - CHFS\$county\*.xlsm" "c:\ricpay\" -force 
    copy-item "\\$inst\RICPAY - CHFS\$county\*.cab" "c:\ricpay\" -force 
    New-Item "c:\Ricpay Temp Folder" -type directory -force
    New-Item "c:\Ricpay Temp Folder\user_profile" -type directory -force
    expand "c:\ricpay\$county.cab" -F:* "c:\Ricpay Temp Folder\user_profile" # NEW For Version 4.23
    expand "c:\ricpay temp folder\user_profile\users.cab" -F:* "c:\Ricpay Temp Folder\user_profile" # NEW For Version 4.23
    copy-item "\\$inst\RICPAY - CHFS\$county\*.tif" "c:\ricpay temp folder\" -Force 
    attrib +R "c:\ricpay temp folder\*.tif" 
    robocopy  "\\$inst\RICPAY - CHFS\Psftp" "c:\program files\psftp" /w:0 /r:0 /mt:128 /NJH /NJS /np 
    foreach ($user in (gci C:\Users).FullName) 
      { 
       if (!($user -like "*public*")) 
         {
          robocopy "c:\Ricpay Temp Folder\user_profile" "$user\AppData\Roaming\ScandAllPro" "ricpay*.*"  /w:0 /r:0 /mt:128 /NJH /NJS /np # NEW For Version 4.23
         }  #added 12/28/2023 - Copies Scanner Profile to all user folders
      } # END Foreach User
   
    # Install code39.ttf Font
    # _____________________________________________________________________________________
    write-host ""
    write-host "Installing CODE 39 Bar Code Font on $env:COMPUTERNAME" -ForegroundColor Magenta
    robocopy "\\$inst\RICPAY - CHFS\$county" "c:\windows\fonts" "code39.ttf" /w:0 /r:0 /mt:128 /NJH /NJS /np
    robocopy "\\$inst\RICPAY - CHFS" "c:\ricpay" "_fontreg.reg" /w:0 /r:0 /mt:128 /NJH /NJS /np
    & regedit /s "c:\ricpay\_fontreg.reg"
    ((Get-Item "HKLM://SOFTWARE\Microsoft\Windows NT\CurrentVersion\fonts").property) -like "CODE 39*"
        
    # Create Shortcut to Header Sheet
    # _____________________________________________________________________________________
    write-host ""
    write-host "Creating MASTERHEADERSHEET Desktop Shortcut on $env:COMPUTERNAME" -ForegroundColor Magenta
    $TargetFile = "c:\ricpay\ricpay"+$county+"master3.xlsm"
    $ShortcutFile = "c:\users\Public\Desktop\MASTERHEADERSHEET.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    copy-item "\\$inst\RICPAY - CHFS\favorites.ico" "c:\ricpay\"
    $Shortcut.IconLocation = "c:\RICPAY\favorites.ico"
    $Shortcut.Save()
   
   # Install TWAIN Scanner Driver
   # _______________________________________________________________________________________________________________
    write-host ""
    write-host "Copying scanner driver files over to $env:COMPUTERNAME" -ForegroundColor Magenta
    robocopy "\\$inst\RICPAY - CHFS\Fujitsu Drivers\Disk1\PSTWAIN\ext" "c:\source\Twain" "psip_twain.msi" /w:0 /r:0 /mt:128 /NJH /NJS /np
    write-host ""
    write-host "Installing TWAIN Scanner driver on $env:COMPUTERNAME" -ForegroundColor Magenta
    Start-Process msiexec.exe -wait -Argumentlist '/i "c:\source\twain\psip_twain.msi" /qn /norestart'
                       
   # Installing Fujitsu ScandAll - Changed this 5/25/21 to make it a little faster....
   # _______________________________________________________________________________________________________________
    write-host ""
    write-host "Installing Fujitsu ScandAll on $env:COMPUTERNAME" -ForegroundColor Magenta
    write-host "Please be patient, depending on network speed, this could take several minutes" -ForegroundColor Magenta
    robocopy "\\$inst\RICPAY - CHFS\ScandAll\ScandAll 2.1.8\SDATA1" "c:\source\ScandAll\ScandAll 2.1.8\SDATA1" "setup.cab" "setup_en.msi" /w:0 /r:0 /mt:128 /NJH /NJS /np
    Start-Process msiexec.exe -wait -Argumentlist '/i "c:\source\ScandAll\ScandAll 2.1.8\SDATA1\setup_en.msi" /qn /norestart'
    Scandall-Rights-Local # Launch Access Rights for Scandall

write-host ""
write-host "RICPAY Installation completed." -ForegroundColor Cyan
pause
exit
  } # End Install Ricpay Local Function

####################################################################################################################
####################################################################################################################

function global:inst-ricpay-remote

  {

    if (!(Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet))

    {
        write-host "$computer offline" -BackgroundColor Red
        break
                
    }

  if (Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet)
 
    {  

        write-host "Installing RicPay on $computer"   -ForegroundColor Magenta
      


    # Copy Ricpay county specific files to computer...
    # _____________________________________________________________________________________
    write-host ""
    write-host "Copying County specific files to $computer" -ForegroundColor Magenta
    New-Item "\\$computer\c$\ricpay" -type directory -force
    copy-item "\\$inst\RICPAY - CHFS\$county\code39.ttf" "\\$computer\c$\ricpay\" -force 
    copy-item "\\$inst\RICPAY - CHFS\$county\*.xlsm" "\\$computer\C$\ricpay\" -force 
    copy-item "\\$inst\RICPAY - CHFS\$county\*.cab" "\\$computer\C$\ricpay\" -force 
    New-Item "\\$computer\c$\ricpay temp folder" -type directory -force
    New-Item "\\$computer\c$\ricpay temp folder\user_profile" -type directory -force
    expand "\\$computer\c$\ricpay\$county.cab" -F:* "\\$computer\c$\ricpay temp folder\user_profile" # NEW For Version 4.23
    expand "\\$computer\c$\ricpay temp folder\user_profile\users.cab" -F:* "\\$computer\c$\ricpay temp folder\user_profile" # NEW For Version 4.23
    copy-item "\\$inst\RICPAY - CHFS\$county\*.tif" "\\$computer\c$\ricpay temp folder\" -Force 
    attrib +R "\\$computer\c$\ricpay temp folder\*.tif"
    robocopy "\\$inst\RICPAY - CHFS\psftp" "\\$computer\c$\program files\psftp" /w:0 /r:0 /mt:128 /NJH /NJS /np

    foreach ($user in (gci \\$computer\c$\Users).FullName) 
     {
      if (!($user -like "*public*"))
        {
         robocopy "\\$computer\c$\Ricpay Temp Folder\user_profile" "$user\AppData\Roaming\ScandAllPro" "RICPAY*.*" /w:0 /r:0 /mt:128 /NJH /NJS /np # NEW For Version 4.23
        }  #added 12/28/2023 - Copies Scanner Profile to all user folders
     }
    
    # Install code39.ttf Font
    # _____________________________________________________________________________________
    write-host ""
    write-host "Installing CODE 39 Bar Code Font on $computer" -ForegroundColor Magenta
    robocopy "\\$inst\RICPAY - CHFS" "\\$computer\c$\windows\fonts"  "code39.ttf" /w:0 /r:0 /mt:128 /NJH /NJS /np
    robocopy "\\$inst\RICPAY - CHFS" "\\$computer\c$\ricpay" "_fontreg.reg" /w:0 /r:0 /mt:128 /NJH /NJS /np
    invoke-command -cn $computer {& regedit /s "c:\ricpay\_fontreg.reg"}
    invoke-command -cn $computer -scriptblock {((Get-Item "HKLM://SOFTWARE\Microsoft\Windows NT\CurrentVersion\fonts").property) -like "CODE 39*"} 
    

    # Create Shortcut to Header Sheet
    # _____________________________________________________________________________________
    write-host ""
    write-host "Creating MASTERHEADERSHEET Desktop Shortcut on $computer" -ForegroundColor Magenta
    $TargetFile = "c:\ricpay\ricpay"+$county+"master3.xlsm"
    $ShortcutFile = "\\$computer\c$\users\Public\Desktop\MASTERHEADERSHEET.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    copy-item "\\$inst\RICPAY - CHFS\favorites.ico" "\\$computer\c$\ricpay\"
    $Shortcut.IconLocation = "c:\ricpay\favorites.ico"
    $Shortcut.Save()
   
   # Install TWAIN Scanner Driver
   # _______________________________________________________________________________________________________________
    write-host ""
    write-host "Copying scanner driver files over to $computer" -ForegroundColor Magenta
    robocopy "\\$inst\RICPAY - CHFS\Fujitsu Drivers\Disk1\PSTWAIN\ext" "\\$computer\c$\source\Twain" "psip_twain.msi" /w:0 /r:0 /mt:128 /NJH /NJS /np 
    write-host ""
    write-host "Installing TWAIN Scanner driver on $computer" -ForegroundColor Magenta
    invoke-command -cn $computer {Start-Process msiexec.exe -wait -Argumentlist '/i "c:\source\Twain\psip_twain.msi" /qn /norestart'}
                       
   # Installing Fujitsu ScandAll - Changed this 5/25/21 to make it a little faster....
   # _______________________________________________________________________________________________________________
    write-host ""
    write-host "Installing Fujitsu ScandAll on $computer" -ForegroundColor Magenta
    write-host "Please be patient, depending on network speed, this could take several minutes" -ForegroundColor Magenta
    new-item -ItemType Directory "\\$computer\c$\source"
    
    robocopy "\\$inst\RICPAY - CHFS\ScandAll\ScandAll 2.1.8\SDATA1" "\\$computer\c$\source\ScandAll\scandall 2.1.8\SDATA1" "setup.cab" "setup_en.msi" /w:0 /r:0 /mt:128 /NJH /NJS /np

    invoke-command -cn $computer {Start-Process msiexec.exe -wait -ArgumentList '/i "c:\source\scandall\ScandAll 2.1.8\SDATA1\setup_en.msi" /qn /norestart'}
    Scandall-Rights-Remote # Launch Access Rights for Scandall                            

write-host ""
write-host "RICPAY Installation completed." -ForegroundColor Cyan
pause;exit} #end test remote computer

} # End Install Ricpay Remote Function

####################################################################################################################
# Main Program Code
####################################################################################################################

function global:county-check

  {
  write-host ""
  $global:county = read-host -prompt "Enter County Name (i.e. Johnson) "

  if ($county -eq "")

    {
     write-warning "Name can't be blank!"
     start-sleep -seconds 2
     Install-Ricpay
    }

  $match=0
  $global:counties = @{}
  $global:counties = import-csv \\$inst\RICPAY - CHFS\counties.csv

 foreach ($global:co in $counties)

  {
    if ($co.cname -eq $county)
    {$match=1
    write-host "!!! County matched !!!" -ForegroundColor Magenta
    write-output "County: $($co.cname)"
    write-output "Number: $($co.cnumber)"
    $code = $co.cnumber
    pause
    return}
  }

 if (($co.cname -ne $county) -and ($co.cnumber -ge "120"))

    {write-warning "Invalid County, or County does not exist!"
     start-sleep -seconds 2
     Install-Ricpay}
} # End County-Check Function

function Global:Scandall-Rights-Local
  {
    write-output "Setting Access to the ScandAll Folder for $env:COMPUTERNAME"
    $identity = 'chfs\domain users'
    $rights = 'Modify' 
    $inheritance = 'ContainerInherit, ObjectInherit' #Other options: [enum]::GetValues('System.Security.AccessControl.Inheritance')
    $propagation = 'None' 
    $type = 'Allow' 
    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($identity,$rights,$inheritance,$propagation, $type)
    $Acl = Get-Acl -Path "c:\ProgramData\ScandAllPRO"
    $Acl.AddAccessRule($ACE)
    Set-Acl -Path "c:\ProgramData\ScandAllPRO" -AclObject $Acl
  }

function Global:Scandall-Rights-Remote
  {
    write-output "Setting Access to the ScandAll Folder for $computer"
    $identity = 'chfs\domain users'
    $rights = 'Modify' 
    $inheritance = 'ContainerInherit, ObjectInherit' 
    $propagation = 'None' 
    $type = 'Allow' 
    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($identity,$rights,$inheritance,$propagation, $type)
    $Acl = Get-Acl -Path "\\$computer\c$\ProgramData\ScandAllPRO"
    $Acl.AddAccessRule($ACE)
    Set-Acl -Path "\\$computer\c$\ProgramData\ScandAllPRO" -AclObject $Acl
  }
 
function global:Install-Ricpay

 {
   $logfile = ""
   #$cred = get-credential
   cls
   write-host ""
   write-host "##### RICPAY Installation Script Version 4.23 #####" -ForegroundColor Magenta
   write-host ""

   #$global:inst = 'eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy'
   # Modified 11/15/24 to use a different repository.
   $global:inst = 'eas.ds.ky.gov\dfs\DS_Share\FS_Software'

   county-check

   $computer = read-host -prompt "Enter Computer Name, Press <ENTER> for Current computer"
   if ($computer -eq "")
    {inst-ricpay-local}
   if ($computer -ne "") 
    {inst-ricpay-remote}

 } # End Install Ricpay Function

# launch Main Program
Install-Ricpay
write-host "Ricpay Install Completed" -ForegroundColor magenta 
pause
# End