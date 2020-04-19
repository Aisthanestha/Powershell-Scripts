<# 
    Author: Greg Hendrickson
    Last Edit:4/19/2020
    
    Updates windows 10 key to BIOS OEM Key

    Requisites:
        Provide Hostname,
        Computer Online,
        WinRM Service Running,
         
    Tools used:
        Invoke-Command
        Get-WmiObject
        slmgr.vbs /dlv (Detailed License)
        slmgr.vbs /ipk (Install Product Key)
#>

Param (
    # Make param mandatory for $ComputerName and offer help message.
    [Parameter(Mandatory=$true, HelpMessage="Provide Name of Remote Computer.")]
    [string]$ComputerName
)

#Receives assigns OEM key and returns results
$Result = Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
   #Gets BIOS Prodcut Key
   $BPK = powershell "(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"; 
   #Assigns BIOS Product Key
   cscript c:/windows/system32/slmgr.vbs -ipk $BPK; 
   $output = cscript c:/windows/system32/slmgr.vbs /dlv; 
   write-host `n`n'OEM Product Key:'"`"$BPK`""`n
   return $output 
}

#Format results

#Could add a function to reduce redundancy.
$ActivationId = $Result | Select-String -Pattern "^Activation ID:\s(.+)"
$ActivationId = $ActivationId.Matches.Groups[1].Value

$LicenseStatus = $Result | Select-String -Pattern "^License Status:\s(.+)"
[string]$LicenseStatus = $LicenseStatus.Matches.Groups[1].Value

$PartialKey = $Result | Select-String -Pattern "^Partial Product Key:\s(.+)"
[string]$PartialKey = $PartialKey.Matches.Groups[1].Value

$PKType = $Result | Select-String -Pattern "^Product Key Channel:\s(.+)"
[string]$PKType = $PKType.Matches.Groups[1].Value

#Final Output
write-host "ActivationID:  "$ActivationId`n"License Status: "$LicenseStatus`n"Partial Key:    "$PartialKey`n"PK Type:        "$PKType`n
