#MoveScan3.0 Made by Charlie and Simon for the Techroom Server.

$jpgcount=0
$pdfcount=0
write-host("This script is monitoring the scans to the server. If found they are transferred from TBU server to Tech Server. Please leave me alone.")
while(1){

$timeday= Get-Date -UFormat %R

    $files = Get-ChildItem z:\* -Include *.jpg, *.pdf | where Length -gt 1kb

    foreach ($file in $files) {
        if ($file -like "*.jpg" ){
            $jpgcount++
            $files = Get-ChildItem C:\Users\chdunsta\Desktop\test\* -Include *.jpg | Move-Item -Destination "D:\Scans\SCAN $jpgcount $(Get-Date -f hh-mm-ss).jpg" -ErrorAction SilentlyContinue

        }
        if ($file -like "*.pdf" ){
            $pdfcount++
            $files = Get-ChildItem C:\Users\chdunsta\Desktop\test\* -Include *.pdf | Move-Item -Destination "D:\Scans\PDF $pdfcount $(Get-Date -f hh-mm-ss).pdf" -ErrorAction SilentlyContinue
        }

        start-sleep -s 1
}


    if($timeday -eq "23:59"){
        $jpgcount=0
        $pdfcount=0
    }

    Start-Sleep -s 3
}