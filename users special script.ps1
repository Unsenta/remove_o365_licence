$proclist = @(
 "Teams",
 "Outlook",
 "winword",
 "Excel",
 "OneDrive",
 "msaccess",
 "mspub",
 "onenote",
 "onenoteim",
 "Powerpnt",
 "lync",
 "Microsoft.AAD.BrokerPlugin"
)

if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64" ) {
  $pfiles = "Program Files"
}
else {
  $pfiles = "Program Files (x86)"
}

function Add-RegistryRecord {
  $registryPath = "HKLM:\Software\EVO\Scripts"
  $Name = "O365LicenceRemoved"
  $value = "1"
  if(!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
  }
  else {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
  }
}

function Remove-OfficeLicense { 
    
    $license = cscript "C:\$pfiles\Microsoft Office\Office16\OSPP.VBS" /dstatus
    $o365 = “---LICENSED---”
    for ($i = 0; $i -lt $license.Length; $i++) {
    #Write-Host $i
        if ($license[$i] -match $o365) {

            $i += 4 #jumping six lines to get to the product key line in the array, check output of dstatus and adjust as needed for the product you are removing
            $keyline = $license[$i] # extra step but i would rather deal with the variable as a string than an array, could be removed i guess, efficiency is not my concern
            $prodkey = $keyline.substring($keyline.length - 5, 5) # getting the last 5 characters of the line (prodkey)
            $check = $prodkey -match '^[0-9A-Z]+$'
            Write-host " $prodkey match 0-9A-Z is $check"
        }
    }
    if ($check -eq $true) {
    cscript "C:\$pfiles\Microsoft Office\Office16\OSPP.VBS" /unpkey:$prodkey
    } else {Write-Host "Key check failed. Exiting"}
    
}


$msgBoxInput =  [System.Windows.MessageBox]::Show('Please do not run any programms for 5 minutes','Office Fix','Ok','Information')
switch  ($msgBoxInput) {
  'Ok' {
    foreach ($proc in $proclist) {
    Write-Host "stopping" $proc
    stop-process -name $proc -ErrorAction SilentlyContinue -Force
    }
    Remove-OfficeLicense
  
  }
}