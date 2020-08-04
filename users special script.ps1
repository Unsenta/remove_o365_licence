Add-Type -AssemblyName PresentationFramework
[reflection.assembly]::loadwithpartialname("system.windows.forms")|Out-Null

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

#if ($env:PROCESSOR_ARCHITECTURE -eq "AMD64" ) {
#  $pfiles = "Program Files"
#}
#else {
#  $pfiles = "Program Files (x86)"
#}

function Add-RegistryRecord {
  $registryPath = "HKLM:\Software\EVO\Scripts"
  if(!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
  }
  else {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
  }
}

function Add-OutlookProfile {
  $profilename = "Evolution"
  $registryPathOutlookProfile = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$profilename"
  $registryPathOutlookDefault = "HKCU:\Software\Microsoft\Office\16.0\Outlook\"
  $value = $profilename
  New-Item -Path $registryPathOutlookProfile -Force | Out-Null
  New-ItemProperty -Path $registryPathOutlookDefault -Name DefaultProfile -Value $value -PropertyType String -Force | Out-Null
  Remove-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover" -name "*@evolutiongaming.com" -Confirm:$false -Force | Out-Null
  
  $Name = "O365ProfileAdded"
  $value = "1"
  Add-RegistryRecord 
}



#function Remove-OfficeLicense { 
#    
#    $license = cscript "C:\$pfiles\Microsoft Office\Office16\OSPP.VBS" /dstatus
#    $o365 = "---LICENSED---"
#    for ($i = 0; $i -lt $license.Length; $i++) {
#    #Write-Host $i
#        if ($license[$i] -match $o365) {
#            $i += 4 #jumping six lines to get to the product key line in the array, check output of dstatus and adjust as needed for the product you are removing
#            $keyline = $license[$i] # extra step but i would rather deal with the variable as a string than an array, could be removed i guess, efficiency is not my concern
#            $prodkey = $keyline.substring($keyline.length - 5, 5) # getting the last 5 characters of the line (prodkey)
#            $check = $prodkey -match '^[0-9A-Z]+$'
#            Write-host " $prodkey match 0-9A-Z is $check"
#        }
#    }
#    if ($check -eq $true) {
#    cscript "C:\$pfiles\Microsoft Office\Office16\OSPP.VBS" /unpkey:$prodkey
#    } else {Write-Host "Key check failed. Exiting"}
#    
#}

function Run-VBScript {
  cscript .\OLicenseCleanup.vbs
  if ($LASTEXITCODE -eq 0) {
    $Name = "O365LicenceRemoved"
    $value = "1"
    Add-RegistryRecord
  }
}



$msgBoxInput =  [System.Windows.MessageBox]::Show('Existing Office installation will be updated to EVOLUTION.COM domain. Switching to new Outlook profile is recommended. 
Switch to new Outlook profile?
 Yes - Switch now (Recommended)
 No - Will make it later','Office update','YesNo','Warning')
 
 
 
 
switch  ($msgBoxInput) {
  'Yes' {
    foreach ($proc in $proclist) {
      Write-Host "stopping" $proc
      stop-process -name $proc -ErrorAction SilentlyContinue -Force
    }
    Run-VBScript
    Add-OutlookProfile
    $Yes = [System.Windows.MessageBox]::Show('Profile created. Office installation updated.','Office update','Ok','Information')  
   }
   'No' {
    foreach ($proc in $proclist) {
      Write-Host "stopping" $proc
      stop-process -name $proc -ErrorAction SilentlyContinue -Force
    }
    Run-VBScript
    $No = [System.Windows.MessageBox]::Show('Profile not created. Office installation updated.','Office update','Ok','Information')  
  }
}