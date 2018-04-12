<#  Script copies an autodiscover file to c:\autodiscovery\ and the registry change 
    file to the user's desktop for the user to run when logged in.
    Designed and created by Phillip Trimble, COT, 2017. Updated 2/26/17  #>

cls

$computer = read-host "Enter destination computer name "

if (!(Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet))

     {
        write-host "$computer offline" -BackgroundColor Red
        break
                
     }

$user = read-host "Enter User name "

New-Item -Path "\\$computer\c$\autodiscovery" -ItemType directory -Force
copy-item -path "\\hfsro121-0581\deploy\autodiscover.xml" "\\$computer\c$\autodiscovery\" -Force
copy-item -path "\\hfsro121-0581\deploy\o365_autodiscover_reg_fix.reg" "\\$computer\c$\users\$user\desktop" -Force


Write-Output ""
Write-Output "Registry file copied to c:\users\$user\desktop, contact user to execute this registry merge."
Write-Output ""
Write-Output "Program complete!"
