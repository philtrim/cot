# Install KYOPS and ScenePD Version for Windows 10 or 11
# COTDSVSEPER043 - DOC049L7VRC573.jus.ds.ky.gov - DOC034L2KXKGY3 - DJJ010LCNTP114
# Last modified: 6/30/25 - COTED2L8BZ6DK3
# $ErrorActionPreference = "SilentlyContinue"

function main-program
{
 cls
 write-host "***** INSTALL KYOPS DUAL OS VERSION *****"
 write-host ""
 $computer = read-host "Enter computer name"
 $credential = Get-Credential -Message "Enter EAS User name and Password"
 Enable-WSManCredSSP -Role Client -DelegateComputer $computer -Force #| out-null     # Fixes Double-Hop Authentication Issue
 Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force } #| out-null # Fixes Double-Hop Authentication Issue
 $session = New-PSSession -cn $computer -Credential $credential -Authentication Credssp

 Invoke-command -session $session {$Win10 = (Get-WmiObject Win32_OperatingSystem).Caption -Match "Windows 10"
       $Win11 = (Get-WmiObject Win32_OperatingSystem).Caption -Match "Windows 11"

 If ($win10)
   {win10}

 If ($win11)
   {win11}

 } # END Invoke-Command Session for Type Of Operating System

} # END Main Program

function win10

{
   
  Invoke-command -session $session -ScriptBlock {write-host "Copying files..." -foreground yellow
  robocopy "\\eas.ds.ky.gov\dfs\ds_share\efs-d2\deploy\jus\KYOPS341" "c:\source\KYOPS341" /s /np /w:0 /r:0 /nfl /ndl /njh /njs}
  clear-install
  Invoke-command -session $session -ScriptBlock {$dotnet = gci "C:\Windows\Microsoft.NET\Framework\*"
  
  if (!($dotnet.name).contains("v3.5"))

    {

      write-warning "DOT NET 3 is not installed"
      write-host "Launching DOTNET 3.5 Installer" -foregroundColor green
      write-host "Loading Image from Windows.ISO File" -foregroundColor green
          
      $ImagePath= "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\COMMON\Fixes-Patches\windows.iso" # Path of ISO image to be mounted
      $ISODrive = (Get-DiskImage -ImagePath $ImagePath | Get-Volume).DriveLetter

      If (!$ISODrive) 
       
        {
          write-host "Mounting Disk Image on $using:computer" -foregroundcolor green
          Mount-DiskImage -ImagePath $ImagePath -StorageType ISO
        }

     $ISODrive = (Get-DiskImage -ImagePath $ImagePath | Get-Volume).DriveLetter
     Write-Host ("ISO Drive is $($ISODrive)")
     write-host "Enabling Feature .NET Framework 3.5 on $using:computer" -foregroundcolor -green
     DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:"$ISODrive":\sources\sxs
     write-host "Dismounting Disk Image on $using:computer" -foregroundcolor -green
     DisMount-DiskImage -ImagePath "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\COMMON\Fixes-Patches\windows.iso"
        
   } # End If DOTNET conains 3.5

   }
 Clear-Install
 Invoke-command -session $session -ScriptBlock {Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops341\scenepd.msidesktopsetup.msi" /qn LID=61476386 LPW=KSP'}
 Invoke-command -session $session -ScriptBlock {Write-Host "Installing ScenePD Active-X Setup" -BackgroundColor DarkCyan}
 Clear-Install
 Invoke-command -session $session -ScriptBlock {Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops341\scenepd.msiactivexsetup.msi" /qn '}
 Clear-Install
 Invoke-command -session $session -ScriptBlock {Write-Host "Installing SQL Server Component for KYOPS" -BackgroundColor DarkCyan}
 Clear-Install
 Invoke-command -session $session -ScriptBlock {Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops341\SqlLocalDB.MSI" /qn IACCEPTSQLLOCALDBLICENSETERMS=YES' }
 Invoke-command -session $session -ScriptBlock {$code = '"C:\source\KYOPS341\setup.exe" /S /v/qn'
 Write-Host "Installing KYOPS Program" -BackgroundColor DarkCyan
 & cmd /c $code}
 Clear-Install
 Invoke-command -session $session -ScriptBlock {$software = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,`
 HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | select DisplayName | where-object {$_.displayname -like "*KYOPS*"}
           
 if ($software)
   {write-host "$($software.displayname) installed successfully." -ForegroundColor YELLOW
    write-host "Program ended" -ForegroundColor YELLOW 
   } #end If

 Disable-WSManCredSSP -Role Server
 Remove-PSSession -Session $session
write-output ""
pause
    } # End Invoke-Command -Session
} # END Function win10

function Clear-Install
        
        {
            Invoke-command -session $session -ScriptBlock {
            $files = ""
            write-host "Function Clear-Install Invoked" -ForegroundColor red
            get-process msiexec | Stop-Process -force -erroraction 'stop'
            $files = gci c:\windows\installer -Recurse | Select-Object fullname,lastwritetime | where {$_.LastWriteTime -gt (date -format d)}
            $files.fullname | remove-item -Recurse -force -Confirm:$false # (this code locates a suspended install and deletes it so as the install will continue)
            get-process msiexec | Start-Process -force # End Invoke Command
         } # END Invoke-Command
        }   # END Function Clear-Install

function win11

   {Invoke-command -session $session -ScriptBlock {write-host "Copying files..." -foreground yellow
    robocopy "\\eas.ds.ky.gov\dfs\ds_share\efs-d2\deploy\jus\KYOPS-NEW" "c:\source\KYOPS-NEW" /s /np /w:0 /r:0} #/nfl /ndl /njh /njs
    clear-install
    Invoke-command -session $session -ScriptBlock {write-host "Installing DotNetfx45" -foreground yellow
    Start-Process -FilePath "C:\source\kyops-new\dotnetfx45_full_x86_x64.exe" -ArgumentList "/q /norestart" -Wait}
    clear-install
    Invoke-command -session $session -ScriptBlock {Write-Host "Installing ScenePD Desktop" -foreground yellow
    Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops-new\spdc5-2-2782.msidesktopsetup.msi" /qn LID=61476386 LPW=KSP'}
    clear-install
    Invoke-command -session $session -ScriptBlock {Write-Host "Installing ScenePD Active-X Setup" -foreground yellow
    Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops-new\spdc5-2-2782.msiactivexsetup.msi" /qn '}
    clear-install
    Invoke-command -session $session -ScriptBlock {Write-Host "Installing SQL Server 2019 Component for KYOPS" -foreground yellow
    Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops-new\SqlLocalDB2019.MSI" /qn IACCEPTSQLLOCALDBLICENSETERMS=YES' }
    clear-install
    Invoke-command -session $session -ScriptBlock {Write-Host "Installing KYOPS-NEW" -foreground yellow
    Start-Process "msiexec.exe" -wait -ArgumentList '/i "c:\source\kyops-new\KYOps.MSI" /qn'}
    Invoke-command -session $session -ScriptBlock {$software = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,`
    HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | select DisplayName | where-object {$_.displayname -like "*KYOPS*"}
         
    if ($software)
     {
      write-host "$($software.displayname) installed successfully." -ForegroundColor YELLOW
      write-host "Program ended" -ForegroundColor YELLOW 
      #pause
     } #end If

    Disable-WSManCredSSP -Role Server
    Remove-PSSession -Session $session
    write-output ""
    pause
    } # End Invoke-Command -Session
 } # END Function Win11

main-program
