# Install new DOR/Revenue User - COTDSVSEPER034 REVSOD725BX04 - REVSOL4SW7473 - PLT 4-27-24 -Last Updated 7/30/24
#write-host "Revenue New User Software"


function Install-BlueZone # SELECTION #1
  
  {

    invoke-command -Session $session -ScriptBlock { write-host "Installing Bluezone 64 Bit" -foregroundcolor yellow
      robocopy "$($using:source)\BlueZone\BlueZone Desktop" "c:\source\BlueZone Desktop" /s /w:0 /r:0 /xf 
      Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\BlueZone Desktop\x64\BlueZone.msi" ADDLOCAL=ALL /qn'
      robocopy "$($using:source)\bluezone\Sessions" "C:\users\public\desktop" *.zmd
      robocopy "$($using:source)\bluezone" "C:\Program Files\BlueZone\7.1" bluezone.lic
      if (test-path "C:\users\public\desktop\BlueZone Session manager 7.1 (64-bit).lnk")
        {& cmd /c del "C:\users\public\desktop\BlueZone Session*.lnk"}
      $access = 'icacls "C:\Users\Public\Desktop\*.zmd" /grant "fin\Domain Users":M /t'
      cmd /c $access
            
      } # End-Invoke Command
    pause
    main-menu # Return to Main Menu

  } # End Function Install-Bluezone 

function Install-Java  # SELECTION #2

  {
    
     invoke-command -Session $session -ScriptBlock { write-host "Installing JAVA" -foregroundcolor yellow
     #& cmd /c "$using:source\Java\javasetup8u181.exe" /s 
     & cmd /c "$using:source\Java\jre-8u411-windows-x64.exe" /s 
     robocopy "$using:source\Java" "c:\temp" "exception list.txt"
      } # End Invoke-Command
    pause    
    main-menu # Return to Main Menu

  } # End Install-Java Function

function Install-CAPS  # SELECTION #3

  {
   
   invoke-command -Session $session -ScriptBlock {write-host "Installing CAPS" -foregroundcolor yellow
   robocopy "$using:source\CAPS" "c:\source\CAPS" /s -xd "Version History" -xd "Updates"
   robocopy "$using:source\CrystalMergeModule" "c:\source\CrystalMergeModule" /s 
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\CAPS\current\CAPS.MSI" /qn' 
   $ShortcutTarget= "C:\Program Files (x86)\CAPS\CAPS.exe"
   $ShortcutFile = "c:\users\public\desktop\CAPS.lnk"
   $WScriptShell = New-Object -ComObject WScript.Shell
   $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
   $Shortcut.TargetPath = $ShortcutTarget
   $Shortcut.Save()
   robocopy "$using:source\CAPS" "c:\source\CAPS" "CAPS-202-ODBC.reg"
   cmd /c regedit /s "c:\source\caps\CAPS-202-ODBC.reg"
   } # End Invoke-Command
   pause    
   main-menu # Return to Main Menu
    Start-Menu

   } # END Function Install-CAPS


function Install-Maillog     # SELECTION #4
  
  {

   invoke-command -Session $session -ScriptBlock {write-host "Installing Maillog" -foregroundcolor yellow
   robocopy "$using:source\MailLog\5_V1.05" "c:\source\MailLog\5_V1.05" /s
   #robocopy "$using:source\CAPS" "c:\source\CAPS" /s
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\MailLog\5_V1.05\mail log system.msi" /qn' 
   robocopy "c:\program files (x86)\Mail Log System" "c:\program files\Mail Log System" /s 
   $ShortcutTarget= "C:\Program Files (x86)\Mail Log System\Maillog.exe"
   $ShortcutFile = "c:\users\public\desktop\Maillog.lnk"
   $WScriptShell = New-Object -ComObject WScript.Shell
   $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
   $Shortcut.TargetPath = $ShortcutTarget
   $Shortcut.Save()
   robocopy "$using:source\MailLog" "c:\source\MailLog" "MAILLOG-301-ODBC.reg"
   cmd /c regedit /s "c:\source\MailLog\MAILLOG-301-ODBC.reg"
   
   
   } # End Invoke-Command
   pause
   main-menu # Return to Main Menu

  } # End Install-MailLog Function


function Install-Oracle11g    # SELECTION #5
  
  {   
    
     invoke-command -Session $session -ScriptBlock { write-host "Installing Oracle 11.1" -foregroundcolor yellow
     write-host "Remote into computer and install Oracle 11.1g from here, \\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\DOR\Oracle11.1g\setup.exe"
     write-host "Also make sure to copy the TNSNAME.ORA from here to here \\$computer\c$\Oracle\apps\product\11.1.0\client\network\admin"

     #robocopy "$using:source\Oracle11.1g" "c:\source\Oracle11.1g" /s 
     #& cmd /c c:\source\Oracle11.1g\oracle-install-response-dor.cmd
          
      } # End Invoke-Command

   pause
   main-menu # Return to Main Menu

  } # End Install-Oracle11g Function

function Install-Visual-Studio    # SELECTION #6
  
  {

   robocopy "\\eas.ds.ky.gov\dfs\ds_share\EFS-D2\Deploy\DOR" "\\$computer\c$\source" "Microsoft Visual Studio 6.0 Enterprise Edition.zip"
   #robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor" "\\$computer\c$\source" MSVSEE6.zip
   Invoke-command -Session $session -ScriptBlock {write-host "Installing Visual Studio 6 Components, this may take a few minutes..." -foregroundcolor yellow
   #Expand-Archive "c:\source\MSVSEE6.zip" -DestinationPath "c:\source" -Force
   Expand-Archive "c:\source\Microsoft Visual Studio 6.0 Enterprise Edition.zip" -DestinationPath "c:\source" -Force
   cd "c:\source\Microsoft Visual Studio 6.0 Enterprise Edition\Disk 1\Setup"
   $c = 'start /wait acmsetup.exe /T acmsetup.stf /S "c:\source\Microsoft Visual Studio 6.0 Enterprise Edition\Disk 1" /n "COT" /o "COT" /k 3251779074 /b 1 /gc c:\users\public\vb6_install_log.txt /qtn'
   cmd /c $c
   #cmd /c c:\source\start-install.cmd
    
  } # End Invoke-Command
  
  pause
  main-menu # Return to Main Menu

  } # End Install-Visual Studio Function


function Install-MFE         # SELECTION #7

  {
    invoke-command -Session $session -ScriptBlock {write-host "Copying MFE-III INstaller to Local Drive" -foregroundcolor yellow
    robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\MFE" "c:\source\MFE" /s /njh /njs 
    cmd /c "c:\source\mfe\C++ Batch for MFE.bat"
    robocopy "c:\source\mfe" "c:\users\public\desktop" "*.url"
    cmd /c c:\source\mfe\VC_redist.x86.exe /quiet /norestart
    cmd /c c:\source\mfe\dotNet4.exe /q /norestart
    

     } # End Invoke-Command
  
  pause
  main-menu # Return to Main Menu

  } # End Install MFE Function    

function Install-OneStop     # SELECTION #8
  
  {

    invoke-command -Session $session -ScriptBlock {write-host "Installing ONE-STOP - Copying Shortcut" -foregroundcolor yellow
    robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Onestop" "c:\users\public\desktop" 
   
      } # End Invoke-Command
  
  pause
  main-menu # Return to Main Menu

  } # End Install-OneStop Function

 
function Install-FoxPro-ODBC  # SELECTION #9
  
  {

    invoke-command -Session $session -ScriptBlock {write-host "Installing FoxPro ODBC Driver" -foregroundcolor yellow
    robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Visual FoxPro ODBC Driver" "c:\source\Visual FoxPro ODBC Driver" VFPODBC6.msi
    Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\Visual FoxPro ODBC Driver\vfpodbc6.msi" /qn'
      } # End Invoke-COmmand
  
  pause
  main-menu # Return to Main Menu

  } # End Install-FoxPro-ODBC Function

function Install-Onbase  # SELECTION #10

{
  invoke-command -Session $session -ScriptBlock {write-host "Installing ON-BASE Software" -foregroundcolor yellow
  robocopy "\\eas.ds.ky.gov\dfs\ConfigMgr\Applications\Desktop\Common\OnBase EP5\Thick Client\Thick Client" "c:\source\onbase" /s 
  cmd /c "c:\source\onbase\setup.exe" /q
  #robocopy "c:\source\onbase" "c:\users\public\desktop" *.url 
  
  } # End Invoke-Command  

 pause
 main-menu # Return to Main Menu 

} # End Install-Onbase Function

function Install-Viking  # SELECTION #11

  {
   
   invoke-command -Session $session -ScriptBlock {write-host "Installing Viking" -foregroundcolor yellow
      robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Viking" "c:\source\viking" /s
      & cmd /c "c:\source\viking\Oracle12c\oracle-install-viking.cmd"
      
      } # End Invoke-Command
    pause    
    main-menu # Return to Main Menu
     Start-Menu

   } # END Function Install-Viking

function Install-CACSG  # SELECTION #12

  {
   
   invoke-command -Session $session -ScriptBlock {write-host "Installing CACS-G (Copying Shortcut)" -foregroundcolor yellow
      
      robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\CACS-G" "c:\Source\CACS-G" 
      &cmd /c "c:\Source\CACS-G\jre-8u411-windows-x64.exe" /s
      &cmd /c "c:\Source\CACS-G\jre-8u411-windows-i586.exe" /s
      robocopy "$using:source\CACS-G" "c:\users\public\desktop" *.url
      } # End Invoke-Command
    pause    
    main-menu # Return to Main Menu
     Start-Menu

   } # END Function CACS-G

function Install-CRRuntime  # SELECTION #13

  {
   
   invoke-command -Session $session -ScriptBlock {write-host "Installing Crystal Reports Runtime" -foregroundcolor yellow
      robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Crystal_Reports_Runtime" "c:\source\Crystal_Reports_Runtime" "CRRuntime_64bit_13_0_25.msi"
      Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\Crystal_Reports_Runtime\CRRuntime_64bit_13_0_25.msi" /qn' 
      } # End Invoke-Command
    pause    
    main-menu # Return to Main Menu
     Start-Menu

   } # END Function CRRuntime


function Install-KRCAPPBAR  # SELECTION #14

  {
   
   invoke-command -Session $session -ScriptBlock {write-host "Installing KRC Application Bar & Driver" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\KRC Application Bar & Driver" "c:\source\KRC Application Bar & Driver" /s 
   if (!(test-path c:\temp))
    {new-item -Path c:\temp -ItemType Directory}
   
   cmd /c msiexec.exe /i "C:\source\KRC Application Bar & Driver\Application Bar\V3.1(Oracle & SQL Server)\KRCAppBar.msi" /qn
   cmd /c msiexec.exe /i "C:\source\KRC Application Bar & Driver\Application Driver\V3.0 (Oracle & SQL Server)\KRC Application Manager.msi" /qn
   robocopy "c:\program files (x86)\KRC Applications" "c:\program files\KRC Applications" /s 
   robocopy "c:\source\KRC Application Bar & Driver\W2RECON" "c:\program files (x86)\KRC Applications\W2RECON" /s 
   robocopy "c:\source\KRC Application Bar & Driver\W2RECON" "c:\program files\KRC Applications\W2RECON" /s
   
   #robocopy "$using:source\KRC Application Bar & Driver\manual-copy\w2recon" "c:\program files (x86)\KRC Applications\W2Recon"
   #robocopy "$using:source\KRC Application Bar & Driver\manual-copy\w2recon" "c:\program files\KRC Applications\W2Recon"
   
   $ShortcutTarget= "C:\Program Files (x86)\KRC Applications\KRCAppBar\KRCAppBar.exe"
   $ShortcutFile = "c:\users\public\desktop\KRCAppBar.lnk"
   $WScriptShell = New-Object -ComObject WScript.Shell
   $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
   $Shortcut.TargetPath = $ShortcutTarget
   $Shortcut.Save()

   $paths = 'C:\Program Files\KRC Applications','C:\Program Files (x86)\KRC Applications','c:\Oracle','C:\Program Files (x86)\Oracle','c:\temp'

   foreach ($path in $paths)
     {
      $acl = Get-Acl $path  
      $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Domain Users","Modify", "ContainerInherit,ObjectInherit", "None", "Allow")  
      $acl.SetAccessRule($AccessRule)  
      $acl | Set-Acl $path
      Set-Acl -Path $path $acl
     }

   
   } # End Invoke-Command
   pause    
   main-menu # Return to Main Menu
    Start-Menu

   } # END Function Install-KRCAPPBAR 


function Install-CrystalMerged  # Selection # 15
  
  {
   invoke-command -Session $session -ScriptBlock {write-host "Installing Crystal Merged Module" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\CrystalMergeModule" "c:\source\CrystalMergeModule" /s 
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\CrystalMergeModule\CrystalMergeModule.msi" /qn' 
   } # End Invoke-Command 
   pause
   main-menu # Return to Main Menu
  
  } # END Function Crystal Merged Model

function Install-HCProvider # Selection # 16

  { 
   invoke-command -Session $session -ScriptBlock {write-host "Installing Health Care Provider" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Health Care Provider\Install" "c:\source\Health Care Provider\Install" /s 
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\Health Care Provider\Install\vfpodbc6.msi" /qn'
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\Health Care Provider\Install\Health Care Provider Tax System.msi" /qn'
   cmd /c regedit /s "c:\source\Health Care Provider\Install\ODBC-746.reg"
   cmd /c regedit /s "c:\source\Health Care Provider\Install\ODBC-746-Edit.reg"

    
   } # End Invoke-Command 
   pause
   main-menu # Return to Main Menu

  } # END Function Install-HCProvider

function Install-EFT # Selection # 17

  { 
   Install-Oracle12c
   invoke-command -Session $session -ScriptBlock {write-host "Installing Electronic Funds Transfer" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\EFT\EFT Registration" "c:\source\EFT\EFT Registration" /s 
   Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\EFT\EFT Registration\Electronic Funds Transfer.msi" /qn'
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\DOR\EFT\EFT Registration" "c:\oracle\apps\product\12.1.0\client\network\admin" *.ora
       
   } # End Invoke-Command 
   pause
   main-menu # Return to Main Menu

  } # END Function Install-EFT 

function Install-Oracle12c # Selection # 18

  { 
   invoke-command -Session $session -ScriptBlock {write-host "Installing Oracle 12c" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Oracle12-64" "c:\source\Oracle12-64" /s 
   & cmd /c c:\source\oracle12-64\oracle12-64-silent-install.cmd
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\Oracle12-64\workingTNSNames"  "c:\oracle\apps\product\12.2.0\client\Network\Admin" "*.ora"
       
   } # End Invoke-Command 
   #pause
   #main-menu # Return to Main Menu

  } # END Function Install-Oracle12c


function Install-Protest-Resolution # Selection # 19

  { 
   invoke-command -Session $session -ScriptBlock {write-host "Installing Protest Resolution" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\dor\ProtestRes\V3.2" "c:\source\ProtestRes\V3.2" /s 
   & cmd /c msiexec.exe /i "c:\source\ProtestRes\V3.2\Protest Resolution Program.msi" /qn
   robocopy "c:\source\ProtestRes\V3.2" "C:\Program Files (x86)\Protest Resolution Program" "prp.exe"
   robocopy "c:\source\ProtestRes\V3.2" "c:\users\public\desktop" "prp.exe"
       
   } # End Invoke-Command 
   pause
   main-menu # Return to Main Menu

  } # END Function Install-Protest-Resolution



function Create-Shortcuts # Selection 20

  {

   invoke-command -Session $session -ScriptBlock {write-host "Copying Desktop Shortcuts" -foregroundcolor yellow
   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\DOR\DOR Web Links" "c:\users\public\desktop" /njh /njs /nfl /ndl /np 
   write-host ""
   write-host "Desktop shortcuts created" -foregroundcolor yellow
   } # END Invoke-Command
   pause
   main-menu # Return to Main Menu

  } # END Function Desktop Shortcuts

function Clean-Up

{
 write-host "Closing connections, cleaning up variables." -ForegroundColor yellow
 Invoke-Command -cn $computer {Disable-WSManCredSSP -Role Server}
 Remove-PSSession -Session $session
 write-host "Processing Complete!" -ForegroundColor yellow
}
 
  
function Main-Menu

  {
    cls
    write-host ""
    write-host "*********************** MAIN MENU **********************************" 
    write-host ""
    write-host "1  - BlueZone               11 - Viking"
    write-host "2  - JAVA                   12 - CACS-G"
    write-host "3  - CAPS (prereq 6,9,13)   13 - Crystal Reports Runtime"
    write-host "4  - MailLog (prereq 6,9)   14 - KRC-AppBar & Driver (prereq 5)" 
    write-host "5  - Oracle 11.1            15 - Crystal Merged Module"
    write-host "6  - Visual Studio          16 - Health Care Provider"
    write-host "7  - MFE/MSVC-Redist        17 - EFT (Installs Oracle12c)"
    write-host "8  - One-Stop               18 - Oracle 12c (64it)"
    write-host "9  - FoxPro ODBC Driver     19 - Protest Resolution"
    write-host "10 - ON BASE                20 - Create Desktop Shortcuts"
    write-host ""
    Write-Host "****************** Press ENTER To EXIT *****************************"
    write-host ""


[int]$choice = read-host "Choice"

<#if ($choice -lt 1 -or $choice -gt 20)
    
    { write-warning "Invalid Choice"
     pause    
     main-menu # Return to Main Menu
    }  #>

switch ($choice)
      
    {
      1 {Install-Bluezone}           # Option 1
      2 {Install-Java}               # Option 2
      3 {Install-CAPS}               # Option 3 
      4 {Install-Maillog}            # Option 4
      5 {Install-Oracle11g}          # Option 5
      6 {Install-Visual-Studio}      # Option 6
      7 {Install-MFE}                # Option 7
      8 {Install-OneStop}            # Option 8
      9 {Install-FoxPro-ODBC}        # Option 9
     10 {Install-Onbase}             # Option 10
     11 {Install-Viking}             # Option 11
     12 {Install-CACSG}              # Option 12
     13 {Install-CRRuntime}          # Option 13
     14 {Install-KRCAPPBAR}          # Option 14
     15 {Install-CrystalMerged}      # Option 15
     16 {Install-HCProvider}         # Option 16
     17 {Install-EFT}                # Option 17
     18 {Install-Oracle12c}          # Option 18
     19 {Install-Protest-Resolution} # Option 19
     20 {Create-Shortcuts}           # Option 20 

     Default {Clean-Up}
    }

} # End Function Main-Menu

cls
write-host "**** Revenue New User Setup ****" 
write-host ""
$global:computer = read-host "Computer"

if (!(test-path "\\$computer\c$"))

 {
  write-warning "$computer is not responding!"
  pause
  Main-Menu
 }

#$global:source = '\\kyprdesxminos\krcapps'
$global:source = '\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\DOR'
$global:cred = Get-Credential 'eas\e-phillip.trimble'
write-host "Setting up secure connection with remote computer" -ForegroundColor yellow
Enable-WSManCredSSP -Role Client -DelegateComputer $computer -Force # Fixes Double-Hop Authentication Issue (Allows this computer to send creds to target computer)
Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force }  # Fixes Double-Hop Authentication Issue (Allows target computer to receive and share creds)
$global:session = New-PSSession -cn $computer -Credential $cred -Authentication Credssp

main-menu