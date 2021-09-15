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

function Get-ApplicationObjectFromServer($appName,$siteServerName)
{


    $path = [string]::Format("\\{0}\ROOT\sms\site_{1}", $siteServerName, $siteCode)
    $scope = new-object System.Management.ManagementScope -ArgumentList $path

    $query = [string]::Format("select * from sms_application where LocalizedDisplayName='{0}' AND ISLatest='true'", $appName.Trim())

    $oQuery = new-object System.Management.ObjectQuery -ArgumentList $query
    $obectSearcher = new-object System.Management.ManagementObjectSearcher -ArgumentList $scope,$oQuery
    $applicationFoundInCollection = $obectSearcher.Get()
    $applicationFoundInCollectionEnumerator = $applicationFoundInCollection.GetEnumerator()

    if($applicationFoundInCollectionEnumerator.MoveNext())
    {
        $returnValue = $applicationFoundInCollectionEnumerator.Current
        $getResult = $returnValue.Get()
        $sdmPackageXml = $returnValue.Properties["SDMPackageXML"].Value.ToString()
        [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($sdmPackageXml)
    }
}

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

#------------------------SETTING CONTENT LOCATION FOR APPS---------------------------------------------------

#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "SCCM 2012 APPLICATION IMPORT  SETTING CONTENT LOCATION              by Abraham Mathew"
$objForm.Size = New-Object System.Drawing.Size(700,450) 
$objForm.StartPosition = "CenterScreen"


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
        $ConsoleFolder=$objTextBox2.Text
    }
    Else # Invalid
    {
             [windows.forms.messagebox]::show($objLabel1.Text,"Enter Input")
    }
    If($objTextBox.Text.Length -gt 0) # Valid
    {
        $VendorFolderLocation=$objTextBox.Text
    }
    Else # Invalid
    {
        [windows.forms.messagebox]::show($objLabel3.Text,"Enter Input")
    }
 
    If($objTextBox2.Text -ne "" -and  $objTextBox.Text -ne "" ) # Valid
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
$objLabel.Location = New-Object System.Drawing.Size(100,60)
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
$objLabel.Text = "Path to CSV file"
$objForm.Controls.Add($objLabel) 
$objLabel.font = $LabelFont

$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Size(100,175) 
$objLabel3.Size = New-Object System.Drawing.Size(280,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel3.Text = "CSV File location"
$objLabel3.font = $LabelFont
$objForm.Controls.Add($objLabel3) 




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


$CSV_FILE=$objTextBox.Text
$SiteServer=$objTextBox2.Text

    If($CSV_FILE -eq "" -or  $SiteServert -eq "" ) # Valid
    {
        EXIT(0)
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
        "IMPORT MODULE FAILED" | out-file  C:\Windows\Logs\CONTENT.log -append
            }

$Setsite = $SiteCode + ":"
CD $Setsite
#$CSV_FILE="C:\Users\administrator.BABULAB\Desktop\Content-Setting.csv" 
$CSV=Import-csv $CSV_FILE | foreach {
$APP_VAR =[String]$_.ServerPath
$ProductionServerPath= $APP_VAR.trim(" ")

$APP_VAR =[String]$_.ApplicationName
$ApplicationName =$APP_VAR.trim(" ")

$APP_VAR =[String]$_.DeploymentTypeNo
$DEPLOYMENT_TYPE_NO= $APP_VAR.trim(" ")

"ApplicationName :" + $ApplicationName + " DeploymentTypeNumber : " + $DEPLOYMENT_TYPE_NO | out-file  C:\Windows\Logs\CONTENT.log -append

$application = [wmi](Get-WmiObject SMS_Application -Namespace root\sms\site_$($SiteCode) |  where {($_.LocalizedDisplayName -eq "$($ApplicationName)") -and ($_.IsLatest)}).__PATH

$applicationXML = Get-ApplicationObjectFromServer "$($ApplicationName)" $SiteServer
$I = 1
foreach ($a in $applicationXML.DeploymentTypes)
            {

               IF($DEPLOYMENT_TYPE_NO -EQ $I){ 
                $a.Installer.Contents[0].Location = $ProductionServerPath
}
$I=$I + 1
                
            }
$newappxml = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::Serialize($applicationXML, $false)

$application.SDMPackageXML = $newappxml
$application.Put() | Out-Null 
}