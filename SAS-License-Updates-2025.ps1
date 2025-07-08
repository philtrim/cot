# Last updated 7/7/25 for July 2025 License, expires 30JUN2026
# KFWF7L6D5C473 - SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt, SAS94_9D2CCC_70006657_Win_Wrkstn.txt
# .\psexec.exe \\$computer "$instloc\SASRenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9CFG7S_70006657_Win_Wrkstn.txt" -erroraction silentlycontinue
cls
write-output "**** SAS License Update ****"
#$computers = get-content "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\SAS\not-updated.txt"
$computers = read-host "Computer"

foreach ($computer in $computers)

{
   if (test-path \\$computer\c$)
    
    {

      $sasinstalled = Get-WmiObject -Class Win32_Product -Computer $computer | select-object @{Name="ComputerName";Expression={$computer}}, `
      Name, Version, InstallDate | where-object {$_.Name -like "SAS*"} 
     
      if ($sasinstalled)

        {
          write-host "SAS Install found" -foregroundcolor yellow
          #start-sleep -seconds 3
          $os = (Get-WmiObject Win32_OperatingSystem -computername $computer).OSArchitecture
   
          if ($os -like "*64*")

            {
              write-host "$computer is 64bit" -foregroundcolor yellow
              start-sleep -seconds 3
             
              $global:sasrenew = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASRenew.exe" -recurse -erroraction silentlycontinue}
              $global:sasdm = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASDM.exe" -recurse -erroraction silentlycontinue}

              
              if ($sasrenew)
           
                {$global:instfolder=$sasrenew.directory}
                
              if ($sasdm)
                
                 {
                   if ($sasdm.length -gt 1)
                    
                    {$sasdm = $sasdm[0]}
               
                 $global:instfolder=$sasdm.directory
              
                }

              $global:licenseloc = (invoke-command -cn $computer {get-childitem -path "c:\Program Files\sas*\licenses" -Filter "*.txt" -recurse})
             
               if (!($licenseloc))
                {
                 write-warning "Can't find license file"
                 #start-sleep -Seconds 3
                 $date = get-date -Format "yyyy-MM-dd-mm-ss"
                 $computer | out-file "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\sas\needs-manual-install-$date.txt"
                 return
                 }
               else
                {
                $global:licenseloc = $licenseloc.Directory             
                #$global:instloc=$instloc.directory
                robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\SAS\64\SASRenewalUtility\9.4" "\\$computer\c$\Program Files\sashome\licenses" "SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt"
                #invoke-command -cn $computer {cmd /c "$instloc\SASDMR.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt" }
          
                

                if ($sasrenew)
                 { 
                  invoke-command -cn $computer {cmd /c "c:\Program Files\SASHome\SASFoundation\9.4\SASRenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt"}
                  #.\psexec.exe \\$computer "$instfolder\sasrenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt" -erroraction silentlycontinue}
                  #$global:launchcmd = '"c:\Program Files\SASHome\SASFoundation\9.4\SASRenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt"'  
                 }


                 
          
                if ($sasdm)
                  {
                                    
                   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\SAS\64\SASRenewalUtility\9.4" "\\$computer\c$\source" "SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt"
                   robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\SAS\64" "\\$computer\c$\source" "sdwresponse.properties"
                   invoke-command -cn $computer {cmd /c "$using:instfolder\SASDM.exe" -quiet -responsefile C:\source\sdwresponse.properties}
                   #.\psexec.exe \\$computer cmd /c c:\Progra~\SASHome\SASFoundation\9.4\sasdm.exe -quiet -responsefile C:\source\sdwresponse.properties
                  }                 
                 
                 #.\psexec.exe \\$computer "$instfolder\SASDMR.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2DH9_70108670_Win_X64_Wrkstn.txt" -erroraction silentlycontinue}
                  
                write-host "Updating license"   -ForegroundColor yellow

                $logfile = Invoke-Command -cn $computer {Get-Childitem "c:\Program Files\*.*" -recurse -filter "setinit.log" -ErrorAction SilentlyContinue} #| where {$_.LastWriteTime -ge ((get-date).date).adddays(-5)}
         
                 if ($logfile)
        
                   {
          
                    $trimlogfile = ($logfile.Directory) -replace (":","$")
                    $logdata = get-content "\\$computer\$trimlogfile\setinit.log" | select-string '2026' -SimpleMatch 
          
                    if ($logdata)
                        {
                        write-output "$computer license log"
                        write-output "----------------------------"
                        $logdata | select-string "expiration"
                        continue
                        }


                    else
                        {
                        
                            write-warning "$computer has invalid log file"
                            write-output "Applying the X86 Version"
                            pause
                            write-output "$computer is 32bit"
                            start-sleep -seconds 3
                            robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\sas\32\SASRenewalUtility\9.4" "\\$computer\c$\program files\sashome\licenses" "SAS94_9D2CCC_70006657_Win_Wrkstn.txt"
            
                            $global:sasrenew = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASRenew.exe" -recurse}
                            $global:sasmdr = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASDMR.exe" -recurse}
            
                            if ($sasrenew)
                
                             {$global:instfolder=$sasrenew.directory
                              .\psexec.exe \\$computer "$instfolder\sasrenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2CCC_70006657_Win_Wrkstn.txt"}

                            if ($sasmdr)
               
                             {$global:instfolder=$sasmdr.directory
                              .\psexec.exe \\$computer "$instfolder\SASDMR.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2CCC_70006657_Win_Wrkstn.txt"}
                             

                            write-output "Rechecking Log File for success"
                            $logdata = get-content "\\$computer\$trimlogfile\setinit.log" | select-string '2025'
          
                                if ($logdata)
                                    {
                                        write-output "$computer log file contains--> $logdata"
                                        continue
                                    }

                                else 
                                        {
                                            write-warning "$computer - log file unable to be updated"
                                            continue
                                        }
                        
                        }              
            
                   } #logfile

    else 
                        {
                        $computer | out-file "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\SAS\no-log-file-found.txt"
                        start-sleep -seconds 3
                        continue
                        } #end chk for updated logfile

                      }
              } #end 64bit section

       else #32bit
        
        {
            write-output "$computer is 32bit"
            start-sleep -seconds 3
            robocopy "\\eas.ds.ky.gov\dfs\DS_Share\EFS-D2\Deploy\sas\32\SASRenewalUtility\9.4" "\\$computer\c$\program files\sashome\licenses" "SAS94_9D2CCC_70006657_Win_Wrkstn.txt"
            
            $global:sasrenew = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASRenew.exe" -recurse}
            $global:sasmdr = invoke-command -cn $computer {get-childitem -path "c:\Program Files" -Filter "SASDMR.exe" -recurse}
            
            if ($sasrenew)
                
                {$global:instfolder=$sasrenew.directory
                  .\psexec.exe \\$computer "$instfolder\sasrenew.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2CCC_70006657_Win_Wrkstn.txt"}

              if ($sasmdr)
               
                {$global:instfolder=$sasmdr.directory
                  .\psexec.exe \\$computer "$instfolder\SASDMR.exe" -s "datafile:c:\program files\sashome\licenses\SAS94_9D2CCC_70006657_Win_Wrkstn.txt"}
            

            $logfile = Invoke-Command -cn $computer {Get-Childitem "c:\Program Files\sashome*\" -recurse -filter "setinit.log"} | where {$_.LastWriteTime -ge ((get-date).date).adddays(-5)}

            if ($logfile)
        
              {
                write-output "$computer has correct and updated log file"
                $gc++
                start-sleep -seconds 3
                continue
              }
            else
              {
               write-warning "$computer has invalid log file, trying an alternative install method.."
               start-sleep -seconds 3
               continue
               }       
               
        } # end 32bit section
        
    } # end SAS installed
         
         else 
            {
             write-warning "SAS Not Installed on $computer"
             start-sleep -seconds 3
             continue
             }


   } #end test path
    else {
          write-warning "$computer not available"
          start-sleep -seconds 3
          continue
         }

} # END Foreach
     