
#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "CHM REPORTING TOOL                                                                                                                                  by Abraham Mathew"
$objForm.Size = New-Object System.Drawing.Size(1100,650) 
$objForm.StartPosition = "CenterScreen"
[bool]$script:Quit = $false

$objForm.BackgroundImageLayout = "Zoom"
    # None, Tile, Center, Stretch, Zoom

$Font = New-Object System.Drawing.Font("Times New Roman",16,[System.Drawing.FontStyle]::Regular)
    # Font styles are: Regular, Bold, Italic, Underline, Strikeout
    
$objForm.Font = $Font


$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(300,320)
$OKButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$OKButton.Text = "Ok"
$OKButton.font = $LabelFont
#$OKButton.Add_Click({$VendorFolderLocation=$objTextBox.Text;$ConsoleFolder=$objTextBox1.Text;$SiteServer=$objTextBox2.Text;$objForm.Close()})
$OKButton.Add_Click({
    If($objTextBox2.Text.Length -gt 0) # Valid
    {
        $CSVfile=$objTextBox2.Text
    }
    Else # Invalid
    {
             [windows.forms.messagebox]::show($objLabel.Text,"Enter Input")
    }
    If($objTextBox.Text.Length -gt 0) # Valid
    {
        $Computerfile=$objTextBox.Text
    }
    Else # Invalid
    {
        [windows.forms.messagebox]::show($objLabel3.Text,"Enter Input")
    }
 
    If($objTextBox2.Text -ne "" -and  $objTextBox.Text -ne "" ) # Valid
    {
        $objForm.Close()
        $script:Quit = $false
    }
    
    })
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(420,320)
$CancelButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$CancelButton.Text = "Cancel"
$CancelButton.font = $LabelFont
$CancelButton.Add_Click({$objForm.Close();$script:Quit = $True})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$LabelFont = New-Object System.Drawing.Font("Times New Roman",20,[System.Drawing.FontStyle]::Bold)
$objLabel.Location = New-Object System.Drawing.Size(100,80)
$objLabel.Size = New-Object System.Drawing.Size(560,80) 
$objLabel.Text = "PLEASE ENTER INPUT INTO THE FORM"
$objLabel.font = $LabelFont


$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(400,175) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(400,285) 
$objLabel.Size = New-Object System.Drawing.Size(400,40)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular) 
$objLabel.Text = "FullPath to CSV file"
$objForm.Controls.Add($objLabel) 
$objLabel.font = $LabelFont

$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Size(100,175) 
$objLabel3.Size = New-Object System.Drawing.Size(280,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel3.Text = "MACHINE DETAILS"
$objLabel3.font = $LabelFont
$objForm.Controls.Add($objLabel3) 




$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(400,250) 
$objTextBox2.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox2)



$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(100,250) 
$objLabel2.Size = New-Object System.Drawing.Size(250,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel2.Text = "CSV FILE LOCATION"
$objForm.Controls.Add($objLabel2)
$objLabel2.font = $LabelFont

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
if($script:Quit -eq 'True'){
exit(0)
}

$Computerfile=$objTextBox.Text
$CSVfile=$objTextBox2.Text

$dir = ([io.fileinfo]$CSVfile).DirectoryName
$UN = $dir + "\Unavailable-Computers.txt"

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

$RemoteComputers = Get-Content $Computerfile
$collectionWithItems = New-Object System.Collections.ArrayList
ForEach ($Computer in $RemoteComputers)
{
Try
{
$item = New-Object System.Object
$p = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ComputerInfo | select cscaption,CsModel,CsNumberOfLogicalProcessors,WindowsVersion,BiosSMBIOSBIOSVersion,CsManufacturer} -ErrorAction Stop
$item | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $p.CsCaption
if($p.CsManufacturer -eq "LENOVO")
{ 
$item | Add-Member -MemberType NoteProperty -Name "Model_Name" -Value (gwmi -ComputerName $Computer -Class Win32_ComputerSystemProduct | select Version).Version
}
else{
$item | Add-Member -MemberType NoteProperty -Name "Model_Name" -Value $p.CsModel
}
$hash = @{}

switch ( $item.Model_Name ){
        default  { $hash.Add('Unknown','Unknown')    }
        "HP EliteBook 850 G5"  { $hash.Add('Kaby Lake R',8)   }        
        "Latitude 7280"        { $hash.Add('Kaby Lake',7)     }        
        "HP EliteBook 840 G3"  { $hash.Add('Skylake',6)       }        
        "HP EliteBook 850 G6"  { $hash.Add('Whiskey Lake',9)  }
        "HP EliteBook 840 G7"  { $hash.Add('Cometlake',10)    }
        
        "ThinkPad T480"  { $hash.Add('Kaby Lake R',8)   }
        "ThinkPad X280"  { $hash.Add('Kaby Lake R',8)   }        
        "ThinkPad T470"  { $hash.Add('Kaby Lake',7)     }        
        "ThinkPad X390"  { $hash.Add('Whiskey Lake',9)  }
        "ThinkPad T590"  { $hash.Add('Whiskey Lake',9)  }        
        "ThinkPad T14G1" { $hash.Add('Cometlake',10)    }
}

$h1 = [String]$hash.Keys
$h2 = [String]$hash.Values
$item | Add-Member -MemberType NoteProperty -Name "Platform" -Value $h1
$item | Add-Member -MemberType NoteProperty -Name "Platform Generation" -Value $h2
             
$item | Add-Member -MemberType NoteProperty -Name "CPU" -Value (gwmi -ComputerName $Computer -Class Win32_Processor | select name).name
$RAM = gwmi Win32_OperatingSystem -ComputerName $Computer | Measure-Object -Property TotalVisibleMemorySize -Sum | % {[Math]::Round($_.sum/1024/1024)}
$item | Add-Member -MemberType NoteProperty -Name "RAM" -Value $RAM
$item | Add-Member -MemberType NoteProperty -Name "Processor Total Logical Cores" -Value $p.CsNumberOfLogicalProcessors
if($p.CsManufacturer -eq "HP"){             
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value ($p.BiosSMBIOSBIOSVersion.Split(" ")[2]).trim(" ")
}elseif($p.CsManufacturer -eq "Dell Inc.")
{
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value $p.BiosSMBIOSBIOSVersion
}
else{
$b = $p.BiosSMBIOSBIOSVersion.Split("(")[1]
$DriverVersion= $b.Split(" ")[0]
$item | Add-Member -MemberType NoteProperty -Name "Bios Version" -Value $DriverVersion
}
$item | Add-Member -MemberType NoteProperty -Name "SMBIOBIOS_Version" -Value $p.BiosSMBIOSBIOSVersion

$item | Add-Member -MemberType NoteProperty -Name "OSShortversion" -Value $p.WindowsVersion
$Version = gwmi win32_product -ComputerName $Computer  -Filter "Name='Google Chrome'" | Select -Expand Version
$item | Add-Member -MemberType NoteProperty -Name "ChromeVersion" -Value $Version

#-----------------------------------------------------WLAN-----------------------------------------------
$g = Invoke-Command -ComputerName $Computer -ScriptBlock {

$DriverVersion = $null
Try { 
If(Test-path "HKLM:\SOFTWARE\Intel\WLAN"){


(gp "HKLM:\SOFTWARE\Intel\WLAN\").WirelessDriverVersion

}
} Catch {} 

}  -ErrorAction Stop 

$item | Add-Member -MemberType NoteProperty -Name "Wireless_Driver_Version" -Value $g

#------------------------------------------------------Bluetooth---------------------------------------------
$b = Invoke-Command -ComputerName $Computer -ScriptBlock {

Try { 
$d = Get-PnpDevice | where{$_.friendlyname -eq "Intel(R) Wireless Bluetooth(R)"}
((Get-PnpDeviceProperty -InstanceId $d.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {} 

}  -ErrorAction Stop  
$item | Add-Member -MemberType NoteProperty -Name "Bluetooth_Driver_Version" -Value $b

#------------------------------------------------------Graphics driver---------------------------------------------
$t = Invoke-Command -ComputerName $Computer -ScriptBlock {

Try { 
$d = Get-PnpDevice -Class "Display" | where{$_.friendlyname -like "Intel(R)*"}
((Get-PnpDeviceProperty -InstanceId $d[0].InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {}  

}  -ErrorAction Stop  

$item | Add-Member -MemberType NoteProperty -Name "Graphics_Driver_Version" -Value $t

#------------------------------------------------------------Audio---------------------------------------------------
$audio = Invoke-Command -ComputerName $Computer -ScriptBlock {

Try { 
$d = Get-PnpDevice | where{$_.friendlyname -eq "Realtek High Definition Audio"}
((Get-PnpDeviceProperty -InstanceId $d.InstanceId | where {$_.keyname -eq "DEVPKEY_Device_DriverVersion"}).data).trim()


} Catch {} 

}  -ErrorAction Stop
$item | Add-Member -MemberType NoteProperty -Name "RealTek_Audio_Driver_Version" -Value $audio

             
$collectionWithItems.Add($item) | Out-Null
}
Catch
{
             Add-Content $UN $Computer 
}
}
$collectionWithItems | Export-Csv $CSVfile -Append
$collectionWithItems






# model format is required
# account for running the tool, check wmi query which is not running
#multiple graphics driver entries













#----Computername
$n = Get-ComputerInfo | select cscaption
#----Computer Model
$n = Get-ComputerInfo | select cscaption
# CPU
$n = Get-ComputerInfo |CsProcessors
#RAM
$n = Get-ComputerInfo |select CsNumberOfLogicalProcessors
#Winver
$n = Get-ComputerInfo |select WindowsVersion

Get-ComputerInfo | out-file "C:\Users\amathe4x\OneDrive - Intel Corporation\PSremoting\1.log"

gwmi -ComputerName 10.106.8.18 -Class Win32_computersystem | select name,model,TotalPhysicalMemory, NumberOfLogicalProcessors
gwmi -ComputerName 10.106.8.18 -Class Win32_Processor | select -Property *
gwmi -ComputerName 10.106.8.18 -Class Win32_OperatingSystem | select version
$session = New-PSSession -ComputerName 10.106.8.18
Invoke-Command -Session $session -ScriptBlock{Get-ComputerInfo | select WindowsVersion }
$Version = gwmi win32_product -ComputerName CQA-t580pac.gar.corp.intel.com  -Filter "Name='Google Chrome'" | Select -Expand Version