param(
[string]$inparameter
)

$startDate=(Get-Date)
If($startDate.Year -gt 2015){
$outhost="Script Expired: Validity is till 2015" | Out-Host
Exit(0)
}
$systemdirectory = Get-ChildItem 'C:\Program Files' | foreach { $_.LastWriteTime.Year -gt 2015} 
$systemdirectory | ForEach-Object {
If($_ -eq "True"){
$outhost="Script Expired: Validity is till 2015" | Out-Host
Exit(0)
}
}


#--------------------------------------------------------------Powershell GUI----------------------------------------------

#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "SCCM 2012 APPLICATION EXPORT                by Abraham Mathew"
$objForm.Size = New-Object System.Drawing.Size(700,450) 
$objForm.StartPosition = "CenterScreen"




#$Image_Dir =get-location
#$Image_Location = $Image_Dir.path +"\Images.jpg"
#$Image = [system.drawing.image]::FromFile($Image_Location)
#$objForm.BackgroundImage = $Image

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
#$OKButton.Add_Click({$ServerPath=$objTextBox.Text;$ConsoleFolder=$objTextBox1.Text;$SiteServer=$objTextBox2.Text;$objForm.Close()})
$OKButton.Add_Click({

    If($objTextBox.Text.Length -gt 0) # Valid
    {
        $ServerPath=$objTextBox.Text
    }
    Else # Invalid
    {
        [windows.forms.messagebox]::show($objLabel3.Text,"Enter Input")
    }
    If($objTextBox2.Text.Length -gt 0) # Valid
    {
        $SiteServer=$objTextBox2.Text
    }
    Else # Invalid
    {
         [windows.forms.messagebox]::show($objLabel2.Text,"Enter Input")
    }
    If($objTextBox2.Text -ne "" -and $objTextBox.Text -ne "" ) # Valid
    {
        $objForm.Close()
    }
    
    })
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(420,320)
$CancelButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$CancelButton.Text = "Cancel"
$CancelButton.font = $LabelFont
$CancelButton.Add_Click({$objForm.Close(); $cancel = $true })
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$LabelFont = New-Object System.Drawing.Font("Times New Roman",20,[System.Drawing.FontStyle]::Bold)
$objLabel.Location = New-Object System.Drawing.Size(100,20)
$objLabel.Size = New-Object System.Drawing.Size(460,30) 
$objLabel.Text = "Please enter the input to the form"
$objLabel.font = $LabelFont


$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(300,175) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(300,205) 
$objLabel.Size = New-Object System.Drawing.Size(400,40)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular) 
$objLabel.Text = "\\Servername\folderlocation\"
$objForm.Controls.Add($objLabel) 
$objLabel.font = $LabelFont

$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Size(100,175) 
$objLabel3.Size = New-Object System.Drawing.Size(280,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel3.Text = "Export File location"
$objLabel3.font = $LabelFont
$objForm.Controls.Add($objLabel3) 

$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(300,100) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox1)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(300,130) 
$objLabel.Size = New-Object System.Drawing.Size(550,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular) 
$objLabel.Text = "For ex:Global Package\SRT-Packages"
$objForm.Controls.Add($objLabel)
$objLabel.font = $LabelFont

$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(100,100) 
$objLabel1.Size = New-Object System.Drawing.Size(250,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel1.Text = "Console Folder"
$objForm.Controls.Add($objLabel1) 
$objLabel1.font = $LabelFont


$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(300,250) 
$objTextBox2.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox2)



$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(100,250) 
$objLabel2.Size = New-Object System.Drawing.Size(250,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel2.Text = "Siteserver"
$objForm.Controls.Add($objLabel2)
$objLabel2.font = $LabelFont

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()


$ServerPath=$objTextBox.Text
$ConsoleFolder=$objTextBox1.Text
$SiteServer=$objTextBox2.Text

    If($ServerPath -eq "" -and $SiteServer -eq "" ) # Valid
    {
        EXIT(0)
    }

#SCCM EXPORT
Function Get-SiteCode($SMSProvider)
{
    $wqlQuery = "SELECT * FROM SMS_ProviderLocation"
    $a = Get-WmiObject -Query $wqlQuery -Namespace "root\sms" -ComputerName $SMSProvider
    $a | ForEach-Object {
        if($_.ProviderForLocalSite)
            {
                $script:SiteCode = $_.SiteCode
            }
    }
return $SiteCode
}
#$SiteServer="Babusccmlab1"
$SiteCode = Get-SiteCode($SiteServer)
Set-Location C:\
If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
try 
            {
                IMPORT-MODULE .\ConfigurationManager.psd1
    }
        catch 
            {      
            $_.Exception.Message
        "IMPORT MODULE FAILED" + $Vendor| out-file  C:\Windows\Logs\EXPORT.log -append
            }

$Setsite = $SiteCode + ":"
CD $Setsite
#$ServerPath = "\\babusccmlab1\Source\"
#$ConsoleFolder="Global Package\SRT-Packages"
If($ConsoleFolder.Length -gt 0) # Valid
    {
       $Confolders=$ConsoleFolder.split("\")
$len=$Confolders.length
$len = $len - 1
try 
            {
                $Parent_Node=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $Confolders[$len]) -and ($_.ObjectType -like '6000') }
    }
        catch 
            {      
            $_.Exception.Message
        "cONSOLE FOLDER NODE FETCH FAILED FAILED" + $Vendor| out-file  C:\Windows\Logs\EXPORT.txt -append
            }

$ParentContainerNodeId = $Parent_Node.ContainerNodeID
}
    
    Else # Invalid
    {
             $ParentContainerNodeId = 0
    }

try 
            {
                $Vendor_folders=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.ParentContainerNodeId -like $ParentContainerNodeId) -and ($_.ObjectType -like '6000') }
    }
        catch 
            {      
            $_.Exception.Message
        "VENDOR FOLDER FETCH FAILED " + $Vendor| out-file  C:\Windows\Logs\EXPORT.log -append
            }


foreach($x in $Vendor_folders.Name)
{

$Vendor_Test =[String]$x

If(!$inparameter -eq ""){
If($inparameter -eq $Vendor_Test){

$Vendor=[String]$x
$ServerPathname = $ServerPath + $Vendor
Set-Location C:\
If(!(Test-Path -path $ServerPathname)) {
try 
            {
                New-item -itemType directory -path $ServerPathname
    }
        catch 
            {      
            $_.Exception.Message
        "FOLDER CREATION FAILED" + $Vendor| out-file  C:\Windows\Logs\EXPORT.log -append
            }
}
If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
$Setsite = $SiteCode + ":"
CD $Setsite
#-----------------------------------------
$FolderObj=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $Vendor) -and ($_.ObjectType -like '6000') }
$FolderObjID = $FolderObj.ContainerNodeID
$ObjectInstance = Get-WmiObject -Class SMS_ObjectContainerItem -ComputerName $SiteServer -Namespace Root\SMS\Site_$SiteCode -filter "ContainerNodeID=$FolderObjID" 
ForEach($Instance in $ObjectInstance) {

$Apps = Get-CMApplication | where { $_.ModelName -eq $Instance.InstanceKey}

foreach ($App in $Apps)
    {
        Export-CMApplication -Path "$(Join-Path $ServerPathname $($App.LocalizedDisplayName)).zip" -ID $($App.CI_ID) -omitcontent
    }
}

}
}
else{

$Vendor =[String]$x
$ServerPathname = $ServerPath + $Vendor
Set-Location C:\
If(!(Test-Path -path $ServerPathname)) {
try 
            {
                New-item -itemType directory -path $ServerPathname
    }
        catch 
            {      
            $_.Exception.Message
        "FOLDER CREATION FAILED" + $Vendor| out-file  C:\Windows\Logs\EXPORT.log -append
            }
}
If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
$Setsite = $SiteCode + ":"
CD $Setsite
#-----------------------------------------
$FolderObj=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $Vendor) -and ($_.ObjectType -like '6000') }
$FolderObjID = $FolderObj.ContainerNodeID
$ObjectInstance = Get-WmiObject -Class SMS_ObjectContainerItem -ComputerName $SiteServer -Namespace Root\SMS\Site_$SiteCode -filter "ContainerNodeID=$FolderObjID" 
ForEach($Instance in $ObjectInstance) {

$Apps = Get-CMApplication | where { $_.ModelName -eq $Instance.InstanceKey}

foreach ($App in $Apps)
    {
        Export-CMApplication -Path "$(Join-Path $ServerPathname $($App.LocalizedDisplayName)).zip" -ID $($App.CI_ID) -omitcontent
    }
}

}


}



