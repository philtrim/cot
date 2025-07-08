# Convert BSYS ZMD files over to ASYS ZMD files
# HFUS037FC8YR74
#$ErrorActionPreference = "SilentyContinue"
cls
write-host "**** Convert BSYS ZMD files over to ASYS ZMD files ****"
write-host ""
$computer = read-host "Computer"
write-host ""
if (!(test-path "\\$computer\c$"))

  {write-warning "$computer is offline";pause;break}

$zmds = gci "\\$computer\c$\program files\bluezone\", "\\$computer\c$\users\*\desktop\" -filter *.zmd -Recurse -Force -erroraction silentlycontinue

foreach ($zmd in $zmds)

  {  
     
    $file = $zmd.fullname
    #$file
    $text = get-content $file | Select-String "host address"
    $version = get-content $file | Select-String "BlueZone Mainframe Display v"
    write-host "$file $text $version" -ForegroundColor yellow
    if ($text -like 'Host Address="cdc2.state.ky.us"')
      
      {write-host "This is a CHFS Mainframe.zmd file" -ForegroundColor green}

    if ($text -like 'Host Address="cdc.state.ky.us"')
      
      {write-host "This is a KYTC/EDU Mainframe.zmd file" -ForegroundColor green}

      write-host ""
    
  
  }