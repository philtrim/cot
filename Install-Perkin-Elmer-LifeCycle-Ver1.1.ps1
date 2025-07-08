# PerkinElmer LifeCycle - 12/20/24
# Tester - COTDSVSEPER015
cls
write-host ""
write-host "**** Install PerkinElmer LifeCycle V1.1 ****"
write-host ""
$computer = read-host "Enter computer"

if (!(test-path "\\$computer\c$"))

  {write-warning "$computer not responding";break}
  
write-host "Setting up secure connection with remote computer" -ForegroundColor yellow
Enable-WSManCredSSP -Role Client -DelegateComputer $computer -Force # Fixes Double-Hop Authentication Issue (Allows this computer to send creds to target computer)
Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force }  # Fixes Double-Hop Authentication Issue (Allows target computer to receive and share creds)
$global:session = New-PSSession -cn $computer -Credential $cred -Authentication Credssp
$global:sgroot = "hfsph121-0372\Public\NBS\PerkinElmer\PROD\SGroot"

#******************************************************************************************************************************************************************************************
# Run setup from the c:\Source\Specimen Gate Laboratory\SGLab -enter this for Specimen Gate config path: \\hfsph121-0372\Public\NBS\PerkinElmer\PROD\SGRoot
# \\hfsph121-0372\Public\NBS\PerkinElmer\PROD\SGRoot
cls 

write-host "**** MANUAL STEP #1 ****" -ForegroundColor green
write-host "**** # STEP 9 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs\SpecimenGate2018\Specimen Gate Laboratory\SGLab" "\\$computer\c$\Source\SpecimenGate2018\Specimen Gate Laboratory\SGLab" /s
write-host "Run setup from c:\Source\SpecimenGate2018\Specimen Gate Laboratory\SGLab" -ForegroundColor Green
write-host "Please make sure to use this PATH in the install location:  \\hfsph121-0372\Public\NBS\PerkinElmer\PROD\SGRoot" -ForegroundColor yellow
pause

# Copy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\COMMON\PerkinElmer Lifecycle\SpecimenGate2018\Specimen History" to "C:\Program Files (x86)\Common Files\PerkinElmer Life Sciences" "WSpeView.ocx" 
write-host "**** # STEP 17 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs\SpecimenGate2018\Specimen History" "\\$computer\c$\Program Files (x86)\Common Files\PerkinElmer Life Sciences" "WSpeView.ocx"
write-host "**** # STEP 18 & 19 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs" "\\$computer\c$\Program Files (x86)\PerkinElmer Life Sciences\Specimen Gate\Bin" "WDisorderUtils.dll" "WQCUtilities.dll"
write-host "**** # STEP 21 IN INSTRUCTIONS ****" -ForegroundColor yellow
invoke-command -Session $session {&cmd /c regsvr32.exe /s  "C:\Program Files (x86)\PerkinElmer Life Sciences\Specimen Gate\Bin\WDisorderUtils.dll"
 &cmd /c regsvr32.exe /s  "C:\Program Files (x86)\PerkinElmer Life Sciences\Specimen Gate\Bin\WQCUtilities.dll"} # END Invoke-Command 
write-host "**** # STEP 22 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs\Shortcuts" "\\$computer\c$\users\public\desktop\shortcuts"
write-host "**** # STEP 23 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs\ResultFlagger" "\\$computer\c$\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger"
write-host "**** # STEP 24 IN INSTRUCTIONS ****" -ForegroundColor yellow
invoke-command -Session $session {&cmd /c regsvr32.exe /s "C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\msvbvm50.dll"
  
 &cmd /c regsvr32.exe /s "C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\Ininreg.dll"
 &cmd /c regsvr32.exe /s "C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\PEResultCalc.dll"
 &cmd /c regsvr32.exe /s "C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\PEResultFlagger.dll"} # END Invoke-Command 

#******************************************************************************************************************************************************************************************
# Run setup from the "c:\Source\Specimen Gate Laboratory\Install v3_1_5"
cls
write-host "**** MANUAL STEP #2 ****" -ForegroundColor green
#write-host "Open PEREsultConfigurator.exe from the C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\" -foregroundcolor green
write-host "**** # STEP 25 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs\LifeCycle- do not run on server\Install v3_1_5" "\\$computer\c$\Source\Specimen Gate Laboratory\Install v3_1_5" /s 
write-host "Run setup from the c:\Source\Specimen Gate Laboratory\Install v3_1_5" -foregroundcolor green
pause

#******************************************************************************************************************************************************************************************
# Run setup from the "c:\Source\Specimen Gate Laboratory\Install v3_1_8"
cls
robocopy "\\$sgroot\installs\LifeCycle- do not run on server\Update v3_1_8" "\\$computer\c$\Source\Specimen Gate Laboratory\Install v3_1_8" /s 
write-host "**** MANUAL STEP #3 ****" -ForegroundColor green
write-host "Run UPDATE.BAT from the c:\Source\Specimen Gate Laboratory\Install v3_1_8" -foregroundcolor green
write-host "The UPDATE.BAT should go through several files" -foregroundcolor yellow
write-host "**** # STEP 30 IN INSTRUCTIONS ****" -ForegroundColor yellow
pause
 
#******************************************************************************************************************************************************************************************
# Install dotnetfx (if needed) from "c:\Source\Specimen Gate Laboratory\RootInstall"
$dotnet = gci "\\$computer\C$\Windows\Microsoft.NET\Framework\*"
if (!$dotnet.name.contains("v3.5"))

      {
       write-warning "DOT NET 3 is not installed"
       write-host "Launching DOTNET 3.5 Installer" -foregroundColor green
       Invoke-Command -Session $session -ScriptBlock {write-host "Inside Session of $using:computer" -foregroundcolor green
         write-host "Enabling Feature .NET Framework 3.5 on $using:computer" -foregroundcolor green
         $Win10 = (Get-WmiObject Win32_OperatingSystem).Caption -Match "Windows 10"
         $Win11 = (Get-WmiObject Win32_OperatingSystem).Caption -Match "Windows 11"
       
         if ($win11)
           {write-host "This is a Windows 11 computer";DISM /online /enable-feature /featurename:NetFX3 /All /LimitAccess /Source:\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\COMMON\Fixes-Patches\win11\sxs}
       
         if ($win10)
           {write-host "This is a Windows 10 computer";DISM /online /enable-feature /featurename:NetFX3 /All /LimitAccess /Source:\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\COMMON\Fixes-Patches\win10\sxs}

         write-host "Disabling WSManCredSSP Role SERVER on $using:computer" -foregroundcolor green
         Disable-WSManCredSSP -Role Server
       } # END INVOKE-COMMAND SESSION

    } # END IF DOTNET not found

#******************************************************************************************************************************************************************************************
# Run setup.exe from "c:\Source\Specimen Gate Laboratory\RootInstall" (Crystal Report Setup)
cls
robocopy "\\$sgroot\installs\LifeCycle- do not run on server\RootInstall" "\\$computer\c$\Source\Specimen Gate Laboratory\RootInstall" /s
write-host "**** MANUAL STEP #4 ****" -ForegroundColor green
write-host "Run setup.exe from c:\Source\Specimen Gate Laboratory\RootInstall (Crystal Report Setup)" -ForegroundColor Green
write-host "**** # STEP 31 IN INSTRUCTIONS ****" -ForegroundColor yellow
pause

robocopy "\\$sgroot\installs\LifeCycle- do not run on server\HotFix" "\\$computer\c$\Source\Specimen Gate Laboratory\HotFix" /s
robocopy "\\$computer\c$\Source\Specimen Gate Laboratory\HotFix\Components" "\\$computer\c$\Program Files (x86)\Common Files\PerkinElmer Life Sciences" /w:0 /r:0

#******************************************************************************************************************************************************************************************
# Run HOTFIX.BAT from c:\Source\Specimen Gate Laboratory\HotFix
cls
write-host "**** MANUAL STEP #5 ****" -ForegroundColor green
write-host "Run HOTFIX.BAT from c:\Source\Specimen Gate Laboratory\HotFix" -ForegroundColor Green
write-host "NOTE: If prompted for location of file or directory always choose directory (D)." -ForegroundColor yellow
write-host "**** # STEP 32 IN INSTRUCTIONS ****" -ForegroundColor yellow
pause

#******************************************************************************************************************************************************************************************
# Run Regasm.bat from c:\Source\Specimen Gate Laboratory\HotFix
cls
write-host "**** MANUAL STEP #6 ****" -ForegroundColor green
write-host "Run Regasm.bat from c:\Source\Specimen Gate Laboratory\HotFix" -ForegroundColor green
write-host "**** # STEP 33 IN INSTRUCTIONS ****" -ForegroundColor yellow
pause

write-host "**** # STEP 34 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs" "\\$computer\c$\Program Files (x86)\Common Files\PerkinElmer Life Sciences" "PEKentuckyDemographics.ocx"
Invoke-command -cn $session {&cmd /c regsvr32.exe /s "C:\Program Files (x86)\PerkinElmer Life Sciences\PEKentuckyDemographics.ocx"
Disable-WSManCredSSP -Role Server}

###  NEW PART For Windows 11?
<#
Make sure all UDLs are pointed to the production server (HFS1VP-SQLS002.chfs.ds.ky.gov,1961). Specifically, the below UDLs:
a.	C:\Program Files (x86)\PerkinElmer Life Sciences\LifeCycle\LifeCycle.UDL
b.	C:\Program Files (x86)\PerkinElmer Life Sciences\LifeCycle\WSecurity.UDL
c.	C:\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger\LifeCycle.UDL
#>
write-host "**** # STEP 35,36,37 IN INSTRUCTIONS ****" -ForegroundColor yellow
robocopy "\\$sgroot\installs" "\\$computer\c$\Program Files\PerkinElmer Life Sciences\LifeCycle" "LifeCycle.UDL" "Live_LifeCycle.UDL" "Test_LifeCycle.UDL" "PerkinElmer LifeCycle.ini"
robocopy "\\$sgroot\installs" "\\$computer\c$\Program Files (x86)\Common Files\PerkinElmer Life Sciences\LifeCycle" "LifeCycle.UDL" "PerkinElmer LifeCycle.ini"
robocopy "\\$sgroot\installs" "\\$computer\c$\Program Files (x86)\PerkinElmer Life Sciences\ResultFlagger" "LifeCycle.UDL"
###  NEW PART

#******************************************************************************************************************************************************************************************
# Add user and give modify rights to the following folder C:\Program Files\Common Files\PerkinElmer Life Sciences# Add user and give modify rights to the following folder C:\Program Files\Common Files\PerkinElmer Life Sciences
cls
write-host "**** STEP #7 ****" -ForegroundColor green
write-host "Granting Modify Access to the appropriate folders" -ForegroundColor Yellow
$paths = 'C:\Program Files\Common Files\PerkinElmer Life Sciences','C:\Program Files\PerkinElmer Life Sciences','C:\Program Files (x86)\PerkinElmer Life Sciences','C:\Program Files (x86)\Common Files\PerkinElmer Life Sciences'

   foreach ($path in $paths)
     {
      $acl = Get-Acl $path  
      $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Domain Users","Modify", "ContainerInherit,ObjectInherit", "None", "Allow")  
      $acl.SetAccessRule($AccessRule)  
      $acl | Set-Acl $path
      Set-Acl -Path $path $acl
     }
#robocopy "\\$installs\ListPro 3.0.38 for Windows 10" "\\$computer\c$\temp\ListPro 3.0.38" 

#******************************************************************************************************************************************************************************************
# Run setup.exe from c:\temp\ListPro 3.0.38 as admin
#cls
#write-host "**** MANUAL STEP #9 ****" -ForegroundColor green
#write-host "Run setup.exe from c:\temp\ListPro 3.0.38 as admin" -foreground green
#pause

write-host "Removing PSESSION on $computer" -foregroundcolor green
   Remove-PSSession -Session $session
   write-output ""
   write-host "Disabling WSManCredSSP Role Client on $computer" -foregroundcolor green
   Disable-WSManCredSSP -Role Client
write-host "";write-host "Program Completed"
#******************************************************************************************************************************************************************************************
#******************************************************************************************************************************************************************************************