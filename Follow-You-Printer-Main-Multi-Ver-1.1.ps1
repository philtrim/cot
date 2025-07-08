# Updated to the latest driver 2/10/25 to support the Altalink C8035
# COTDSVSEPER015 vdi     COT1VP-PSXS001 - COTED2L8BZ6DK3 email
# 5/9/25 - When adding \\server\printer you must install the Printer tool and also restart the print spooler so it will show up
###################################################################################################
function global:computer-online
{
 if (!(test-path "\\$computer\c$"))
   {
    write-warning "Computer is not responding, make sure it is on the network!"
    pause
    install-printer-menu
   }
 #if (test-path "\\$computer\c$")
 #  {continue} 
}
###################################################################################################
function global:open-registry
 {
   write-host "Working with the remote computer" -ForegroundColor green
    $rr = get-service -name "remoteregistry"
    if ($rr.StartType -eq "disabled")
       {set-service $rr -StartType "manual" -PassThru}
    if ($rr.status -ne "Running")
       {Start-service $rr}
 } 


###################################################################################################
function global:printer-admin-tool
{

cls
Write-Output "**** COT Printer Tool ****"
Write-Output ""
$computer = read-host "Enter computer name"
$computer = $computer.trim()
computer-online
Write-Output ""
#[int]$i=0
invoke-command -cn $computer -ScriptBlock {
                    
            $regpath = 'HKLM:\Software\Policies\Microsoft\Windows NT'
            $pfound = gci $regpath | where {$regpath.name -eq "Printers"}
                    
            if (!($pfound))

              { 
               write-host "Creating registry entry for printer tool" -BackgroundColor DarkCyan
               #New-Item -Path "HKLM:\Software\Policies\Microsoft\Windows NT\Printers"  -force
               New-Item -Path "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint" -force 
                       
              }
                  
              write-host "Updating registry entry for printer tool" -BackgroundColor DarkCyan
              Set-ItemProperty -path "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint" -Name "RestrictDriverInstallationToAdministrators" -Value 0 -force
              Get-Item -path "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint"}
              #Start-Sleep -Seconds 5
              pause
              install-printer-menu
} # END Function printer-admin-tool
###################################################################################################
function global:Install-Server-Printer

{
 cls  # /ga add per machine printer connections (the connection will be propagated to the user upon logon)
 write-host "**** Install Follow-You Printers ****"
 $credential = Get-Credential -Message "Enter EAS User Account & Password"
 $computer = read-host "Enter computer"
 $computer = $computer.trim()
 computer-online
 $servername = read-host "Enter Print Server Name"
 $printers = get-printer -cn $servername
 write-host "";write-host "Setting up secure channel with remote devices...." -foregroundcolor yellow
 write-host ""
 Enable-WSManCredSSP -Role Client -DelegateComputer $computer -Force  # Double-Hop
 Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force } # Double-Hop
 $session = New-PSSession -cn $computer -Credential $credential -Authentication Credssp # Double-Hop
 #start-sleep -seconds 10
 invoke-command -Session $session -ScriptBlock {
 
 [int]$global:many = read-host "How Many Printers for this Computer"

 for ($p=0;$p -le $many-1;$p++)
 
 {
  cls
  #write-host "Please be patient, searching $using:servername for available printers.." -ForegroundColor green
  write-host ""
  write-host "Available Printers On $using:servername" -ForegroundColor green
  write-host "***************************************"

  for ($i=0;$i -lt (($using:printers.name).count);$i++)

    {
      write-host $i '=' ($using:printers)[$i].name `r -ForegroundColor yellow
    }

  write-host ""
  [int]$cuser = read-host "Enter Printer # $($p+1) of $many to Install" 

  if ($cuser -ilt 0 -or $cuser -ige $using:printers.name.count)
    
    {
      write-warning "Exiting Program"
      #Start-Sleep -Seconds 3
      write-output "Program Ended"
      clean-up
    }

write-host  "You selected--> $(($using:printers)[$cuser].name)" `n`r -backgroundcolor Darkred

$select = read-host "Continue (y/n)"

if ($select.ToUpper() -eq "Y")

  {
   $command = "\\$($using:servername)\$(($using:printers)[$cuser].name)"
   write-host ""
   write-host "Installing $(($using:printers)[$cuser].name) on $using:computer" -backgroundcolor Darkred
   rundll32 printui.dll,PrintUIEntry /ga /n$command # Add Printer
   #Get-Service -name spooler
   Stop-Service -name spooler -Force
   Start-Service -name spooler
   #Restart-Service "Spooler" -Force # Added 5/9/24 so printer would show up w/o a restart
        
  }

  } # End Iterating thru printers

 } # END Invoke-Command Scriptblock

 pause
 clean-up

} # END Function Install-Server-Printer
###################################################################################################
function global:Install-TCPIP-Printer

{

cls
write-host ""
write-output "**** Install Xerox TCPIP Printer ****"
$global:computers = read-host "Enter Computer(s)"
#computer-online
$global:computers = $computers.split(",")
$global:credential = Get-Credential -Message "Enter EAS User Account & Password"
                
foreach ($computer in $computers)
     
   {
     if (test-path \\$computer\c$)

      {
                  
         #printer-admin-tool # Added 5/9/25
         write-host "";write-host "Setting up secure channel with remote devices...." -foregroundcolor yellow
         write-host ""
         Enable-WSManCredSSP -Role Client -DelegateComputer $computer -force | out-null    # Enables WSMANCredSSP CLIENT Role on LOCAL Computer and adds REMOTE Computer as Delegate
         Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force |out-null} # Enables WSMANCredSSP SERVER Role on REMOTE Computer
         #Start-Sleep -Seconds 5
         $global:session = New-PSSession -cn $computer -Credential $credential -Authentication Credssp   # Creates Session on REMOTE computer specifing -Authentication Credssp
         
         [int]$global:many = read-host "How Many Printers for this Computer"

          for ([int]$i=1;$i -le $many;$i++)
          
          {

           write-host ""
           write-host ""
           $global:printername = read-host "Enter Printer Name # $i of $many"
           $global:ipaddress = read-host "Enter Printer IP Address "
           write-host ""
           write-host ""
           write-host "Installing $printername on $computer" -foreground green
           write-host ""
                                  
           invoke-command -session $session -scriptblock {
         
                    pnputil /add-driver "\\eas.ds.ky.gov\dfs\ds_share\EFS-D2\Deploy\Drivers\Xerox64\*.inf"
                    $inf = Get-ChildItem "\\eas.ds.ky.gov\dfs\ds_share\EFS-D2\Deploy\Drivers\Xerox64"  -filter *.inf
                    
                    if ($inf.count -gt 1)
                      {$inf = $inf[0]}
                    
                    Add-PrinterDriver -name "Xerox Global Print Driver" -InfPath $($inf.fullname)
                    $portexists = Get-PrinterPort | select-object name | where {$_.name -eq "IP_$using:ipaddress"}
               
                if (!($portexists))
                 {
                   add-printerport -name "IP_$using:ipaddress" -PrinterHostAddress "$using:ipaddress"
                 }
                
                Add-Printer -DriverName "Xerox Global Print Driver" -Name "$using:printername" -PortName "IP_$using:IPAddress"
                 
                                         
              } # End Invoke-Command

         } # End For How Many Printers Loop
          
    } # End If computer exists

   if (!(Test-Path \\$computer\c$))

       {
         write-host "$computer offline" -BackgroundColor Red
         write-host "Printer install failed!" -BackgroundColor Red
         #break
         pause
         install-printer-menu
       } # If ! Computer exists
      
     

 } # End Foreach Computer

 write-host ""
 write-host "Installed Printers for $computer"
 $printers = get-printer -ComputerName $computer  | select-object -Property Name, Type, PortName | where-object {$_.Name -NotLike "*Fax*" -and $_.Name -NotLike "*One*" -and $_.name -notlike "*Microsoft*"}
 $printers | out-host
 write-host ""
 Write-Host "Program Ended!" -ForegroundColor green
 write-host ""
 pause
 clean-up

} # END Function Install-TCPIP-Printer
###################################################################################################
Function Global:Get-Remote-Printer

{
 cls
 $ErrorActionPreference =  "silentlycontinue"
 write-host "**** Get Remote Printers ****"
 write-host ""
 $session = $null
 $computer = read-host "Enter Computer"
 computer-online
 $credential = Get-Credential -Message "Enter EAS User Account & Password"

 write-host "";write-host "Setting up secure channel with remote devices...." -foregroundcolor yellow
 write-host ""
 Enable-WSManCredSSP -Role Client -DelegateComputer $computer -Force | out-null     # Fixes Double-Hop Authentication Issue
 Invoke-Command -cn $computer {Enable-WSManCredSSP -Role Server -Force | out-null }  # Fixes Double-Hop Authentication Issue
 $session = New-PSSession -cn $computer -Credential $credential -Authentication Credssp
 Invoke-Command -Session $session -ScriptBlock {
    
    open-registry
    write-host "Pulling user profiles from the registry" -ForegroundColor green
    $users = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -exclude ".default","S-1-5-18","S-1-5-19","S-1-5-20","S*_Classes"
    $usernames = ($users | Get-ItemProperty -name profileimagepath)   # Error was using Get-Item instead of Get-ItemProperty

    for ($u=0;$u -lt ($usernames.Count);$u++)
     
     {
      write-host $u '=' ($usernames[$u].profileimagepath) `r -ForegroundColor yellow
     }

  write-host ""
  [int]$cuser = read-host "Find installed printers for which user#" 
     
  try {(get-item -path Registry::\HKEY_USERS\($usernames[$cuser].pschildname)\Printers\Settings).property}
  catch {write-host ""}
  finally {(get-item -path Registry::\HKEY_USERS\$($usernames[$cuser].pschildname)\Printers\*).property}    

  } # END Invoke-Command
    
    write-host ""
    Stop-service $rr
    set-service $rr -StartType "disabled"
    pause
    clean-up

  }   # END Function Get REmote Printer    
###################################################################################################
function global:clean-up                                   

 {
   write-host"";write-host "Cleaning up...." -foregroundcolor yellow
   Remove-PSSession -Session $session
   write-output ""
   Invoke-Command -cn $computer {Disable-WSManCredSSP -Role Server}
   Disable-WSManCredSSP -Role Client
   install-printer-menu

 } # END Function Clean-Up
###################################################################################################
function global:install-printer-menu
{
 cls
 write-host "****  Install Printer Menu   ****" -foregroundcolor yellow
 write-host "*********************************" -foregroundcolor green
 write-host "* 1 = Run COT Printer Admin Fix *" -foregroundcolor green
 write-host "* 2 = Install \\server\printer  *" -foregroundcolor green
 write-host "* 3 = Install TCP-IP Printer    *" -foregroundcolor green
 write-host "* 4 = Get Installed Printers    *" -foregroundcolor green
 write-host "Press ENTER to END Program      *" -foregroundcolor yellow
 write-host "*********************************" -foregroundcolor green

 $option = read-host "Enter Option"

 switch ($option)
      {
      1 {printer-admin-tool}        # Option 1
      2 {Install-Server-Printer}    # Option 2
      3 {Install-TCPIP-Printer}     # Option 3 
      4 {Get-Remote-Printer}        # Option 4
      Default {break}               # Default

      }
} # END Function Install Printer Menu
###################################################################################################
install-printer-menu