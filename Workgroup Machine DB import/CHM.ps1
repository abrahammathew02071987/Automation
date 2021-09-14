

#----------------------------------------------------------------------------------------------------------------------------------------------
Function check-validity
{
$startDate=(Get-Date)
If($startDate.Year -gt 2021){

Exit(0)
}
$systemdirectory = Get-ChildItem 'C:\Program Files' | foreach { $_.LastWriteTime.Year -gt 2021} 
$systemdirectory | ForEach-Object {
If($_ -eq "True"){

Exit(0)
}
}
}
check-validity




$item = New-Object System.Object
$p = Get-ComputerInfo | select cscaption,CsModel,CsNumberOfLogicalProcessors,WindowsVersion,BiosSMBIOSBIOSVersion,CsManufacturer,CsProcessors
$item | Add-Member -MemberType NoteProperty -Name "ID" -Value 1
$item | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $p.CsCaption
if($p.CsManufacturer -eq "Lenovo"){
$item | Add-Member -MemberType NoteProperty -Name "Model_Name" -Value (Get-WMIObject -class Win32_ComputerSystem | select systemfamily).systemfamily
}else{
$item | Add-Member -MemberType NoteProperty -Name "Model_Name" -Value $p.CsModel
}
$hash = @{}

switch ( $item.Model_Name ){
        default  { $hash.Add('Unknown','Unknown')    }
        "HP EliteBook 850 G5"  { $hash.Add('Kaby Lake R',8)   }        
        "Latitude 7280"        { $hash.Add('Kaby Lake',7)     }        
        "HP EliteBook 840 G3"  { $hash.Add('Skylake',6)       }        
        "HP EliteBook 850 G6"  { $hash.Add('Whiskey Lake',9)  }
        "HP EliteBook 840 G7 Notebook PC"  { $hash.Add('Comet Lake',10)    }
        
        "ThinkPad T480"  { $hash.Add('Kaby Lake R',8)   }
        "ThinkPad X280"  { $hash.Add('Kaby Lake R',8)   }        
        "ThinkPad T470"  { $hash.Add('Kaby Lake',7)     }        
        "ThinkPad X390"  { $hash.Add('Whiskey Lake',9)  }
        "ThinkPad T590"  { $hash.Add('Whiskey Lake',9)  }        
        "ThinkPad T14 Gen 1" { $hash.Add('Comet Lake',10)    }
}

$h1 = [String]$hash.Keys
$h2 = [String]$hash.Values
$item | Add-Member -MemberType NoteProperty -Name "Platform" -Value $h1
$item | Add-Member -MemberType NoteProperty -Name "Platform Generation" -Value $h2
             
$item | Add-Member -MemberType NoteProperty -Name "CPU" -Value (gwmi -Class Win32_Processor | select name).name
$RAM = gwmi Win32_OperatingSystem  | Measure-Object -Property TotalVisibleMemorySize -Sum | % {[Math]::Round($_.sum/1024/1024)}
$item | Add-Member -MemberType NoteProperty -Name "RAM" -Value $RAM
$item | Add-Member -MemberType NoteProperty -Name "Processor Total Logical Cores" -Value $p.CsNumberOfLogicalProcessors
if($p.CsManufacturer -eq "HP"){             
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value ($p.BiosSMBIOSBIOSVersion.Split(" ")[2]).trim(" ")
}elseif($p.CsManufacturer -eq "Dell Inc.")
{
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value $p.BiosSMBIOSBIOSVersion
}else{
$b = $p.BiosSMBIOSBIOSVersion.Split("(")[1]
$DriverVersion= $b.Split(" ")[0]
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value $DriverVersion
}
$item | Add-Member -MemberType NoteProperty -Name "SMBIOBIOS_Version" -Value $p.BiosSMBIOSBIOSVersion

$Version = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'

$item | Add-Member -MemberType NoteProperty -Name "OS" -Value "Version $($Version.ReleaseId) (OS Build $($Version.CurrentBuildNumber).$($Version.UBR))"

$item | Add-Member -MemberType NoteProperty -Name "ChromeVersion" -Value ((Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo).ProductVersion

#-----------------------------------------------------WLAN-----------------------------------------------
$Istrue = $false

$DriverVersion = $null
Try { 
$IMEI = Get-PnpDevice -Class NET | where{$_.friendlyname -like "Intel(R) Wi-Fi 6*"}
$DriverVersion = ((Get-PnpDeviceProperty -InstanceId $IMEI.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {}
if($DriverVersion){
$item | Add-Member -MemberType NoteProperty -Name "Wireless_Driver_Version" -Value $DriverVersion
$Istrue = $True
}
$DriverVersion = $null
Try { 
$IMEI = Get-PnpDevice -Class NET | where{$_.friendlyname -like "Intel(R) Wireless-AC*"}
$DriverVersion = ((Get-PnpDeviceProperty -InstanceId $IMEI.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {}
if($DriverVersion){
$item | Add-Member -MemberType NoteProperty -Name "Wireless_Driver_Version" -Value $DriverVersion
$Istrue = $True
}
$DriverVersion = $null
Try { 
$IMEI = Get-PnpDevice -Class NET | where{$_.friendlyname -like "Intel(R) Dual Band Wireless-AC*"}
$DriverVersion = ((Get-PnpDeviceProperty -InstanceId $IMEI.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {}
if($DriverVersion){
$item | Add-Member -MemberType NoteProperty -Name "Wireless_Driver_Version" -Value $DriverVersion
$Istrue = $True
}
if($Istrue -eq $false){
$item | Add-Member -MemberType NoteProperty -Name "Wireless_Driver_Version" -Value "Unknown"
}





#------------------------------------------------------Bluetooth---------------------------------------------
$DriverVersion = $null

Try { 
$d = Get-PnpDevice | where{$_.friendlyname -eq "Intel(R) Wireless Bluetooth(R)"}
$DriverVersion = ((Get-PnpDeviceProperty -InstanceId $d.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {} 

 
$item | Add-Member -MemberType NoteProperty -Name "Bluetooth_Driver_Version" -Value $DriverVersion

#------------------------------------------------------Graphics driver---------------------------------------------
$DriverVersion = $null

Try { 
$d = Get-PnpDevice -Class "Display" | where{$_.friendlyname -like "Intel(R)*"}
$DriverVersion =((Get-PnpDeviceProperty -InstanceId $d[0].InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {}  

 

$item | Add-Member -MemberType NoteProperty -Name "Graphics_Driver_Version" -Value $DriverVersion

#------------------------------------------------------------Audio---------------------------------------------------
$DriverVersion = $null
$Audiodrivers = @()

Try { 

#Get-PnpDevice -Class media | where{$_.Status -eq "OK"} | % {$Audiodrivers.Add($_.FriendlyName,((Get-PnpDeviceProperty -InstanceId $_.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim())}
Get-PnpDevice -Class media | where{$_.Status -eq "OK"} | % {
$DriverVersion = (Get-PnpDeviceProperty -InstanceId $_.InstanceId | where{$_.keyname -eq "DEVPKEY_Device_DriverVersion" }).data 
$audio = $_.FriendlyName + "=" +$DriverVersion
$Audiodrivers += ,$audio
}


} Catch {} 

$Audiodrivers = $Audiodrivers | Out-String
$item | Add-Member -MemberType NoteProperty -Name "RealTek_Audio_Driver_Version" -Value $Audiodrivers

             




$item | Export-Csv C:\EnterpriseTestbed\Result\CHM_WMI.csv 
$item






# model format is required
# account for running the tool, check wmi query which is not running
#multiple graphics driver entries












