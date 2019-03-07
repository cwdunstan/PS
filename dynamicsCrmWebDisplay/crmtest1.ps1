#REQUIRED TO ESTABLISH CONNECTION
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

#GET USER CREDENTIALS
$Cred = Get-Credential -Message "Please sign in with your Microsoft Account:"

#Establish CRM Connection
if($CRMconn -eq $null)
{
$CRMconn = Get-CrmConnection -ServerUrl https://msretailstore.crm.dynamics.com -Credential $Cred -OrganizationName msretailstore
}


#GET ACTIVE BOH CASES
function Get-Appointments-boh
{
    param ([System.Object]$CRMConnection)
    $temp = Get-CrmRecordsByViewName -conn $CRMConnection -ViewName "Cases - BOH Service" 
    Write-Host $temp.Count "BoH Cases Loaded from CRM."
    write-Host ("")
    return $temp
}
$bohAppointments =  Get-Appointments-boh -CRMConnection $CRMconn
$bohAppCount = $bohAppointments.Count

#boh APPOINTMENT DETAILS
function boh-details
{
    param ([System.Object]$bohAppointments)
    $i=0
    $bohstore= @()
    while($i -lt $bohAppCount){
    $boh = New-Object PSObject -property @{
            User= $bohAppointments.CrmRecords[$i].customerid
            Bin= $bohAppointments.CrmRecords[$i].ms_binnumberassigned
            Case= $bohAppointments.CrmRecords[$i].ticketnumber
            Notes= $bohAppointments.CrmRecords[$i].ms_reasonforcase
        }
        $bohstore += $boh
        $i++
    }
    return $bohstore
}

#GET RFP CASES
function Get-Appointments-rfp
{
    param ([System.Object]$CRMConnection)
    $temp = Get-CrmRecordsByViewName -conn $CRMConnection -ViewName "Cases - Ready For Pickup" 
    Write-Host $temp.Count "RFP Cases Loaded from CRM."
    write-Host ("")
    return $temp
}
$rfpAppointments =  Get-Appointments-rfp -CRMConnection $CRMconn
$rfpAppCount = $rfpAppointments.Count

#rfp APPOINTMENT DETAILS
function rfp-details
{
    param ([System.Object]$rfpAppointments)
    $i=0
    $rfpstore= @()
    while($i -lt $rfpAppCount){
    $rfp = New-Object PSObject -property @{
            User= $rfpAppointments.CrmRecords[$i].customerid
            Bin= $rfpAppointments.CrmRecords[$i].ms_binnumberassigned
            Case= $rfpAppointments.CrmRecords[$i].ticketnumber
            Contacted= $rfpAppointments.CrmRecords[$i].ms_lastcontactedon
        }
        $rfpstore += $rfp
        $i++
    }
    return $rfpstore
}


$bohhead = @"
<div id="h1wrap"><h1> Back of House </h1></div>
<h2>BoH Devices</h2>
"@
$Header = @"
<link href="main.css" rel="stylesheet" type="text/css">
"@
$rfphead = @"
<h2>RFP Devices</h2>
"@

$bohhead | Out-File "C:\Users\chdunsta\OneDrive - Microsoft\PSjects\html\BoHomepage.html"
$bohprint = boh-details $bohAppointments | ConvertTo-Html -Property Bin,User,Case,Notes -Head $Header  | Add-Content "C:\Users\chdunsta\OneDrive - Microsoft\PSjects\html\BoHomepage.html"
$rfphead | Add-content "C:\Users\chdunsta\OneDrive - Microsoft\PSjects\html\BoHomepage.html"
$rfpprint = rfp-details $rfpAppointments | ConvertTo-Html -Property Bin,User,Case,Contacted  | Add-content "C:\Users\chdunsta\OneDrive - Microsoft\PSjects\html\BoHomepage.html"




