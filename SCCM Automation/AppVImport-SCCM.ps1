$startDate=(Get-Date)
If($startDate.Year -gt 2017){
$outhost="Script Expired: Validity is till 2017" | Out-Host
Exit(0)
}
$systemdirectory = Get-ChildItem 'C:\Program Files' | foreach { $_.LastWriteTime.Year -gt 2017} 
$systemdirectory | ForEach-Object {
If($_ -eq "True"){
$outhost="Script Expired: Validity is till 2017" | Out-Host
Exit(0)
}
}

#------------------------SETTING CONTENT LOCATION FOR APPS---------------------------------------------------

#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "SCCM 2012 APPLICATION IMPORT  TO CONSOLE             by Abraham Mathew"
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
$objLabel.Text = "FullPath to CSV file"
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


    If($CSV_FILE -eq "" -or $SITESERVER -eq "" ) # Valid
    {
        EXIT(0)
    }


#----------------------------------------------------Move-CMObject------------------------------------------------------------
 
#Function to Move Object.
#When you import the application it would be under root directory.Move function is used to move to corresponding Vendor folder.
 
 
Function Move-CMObject
{
    [CmdLetBinding()]
    Param(
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Site code")]
              $SiteCode,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Name")]
              $SiteServer,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Object ID")]
              [ARRAY]$ObjectID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter current folder ID")]
              [uint32]$CurrentFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter target folder ID")]
              [uint32]$TargetFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Ente .r obje .ct type ID")]
              [uint32]$ObjectTypeID              
        )
 Try{
    
        Invoke-WmiMethod -Namespace "Root\SMS\Site_$SiteCode" -Class SMS_objectContainerItem -Name MoveMembers -ArgumentList $CurrentFolderID,$ObjectID,$ObjectTypeID,$TargetFolderID -ComputerName $SiteServer -ErrorAction STOP
    }
    Catch{
        $date = get-date     
            $_.Exception.Message
        $date , "IMPORT MODULE FAILED - Move"| out-file  C:\Temp\IMPORT.log -append
    }   
 
}
#-----------------------------------------------------New-CMFolder-------------------------------------------------------------------
#Function for Creating the Folder
#This Function is used for console folder creation
Function New-CMFolder
{
    [CmdLetBinding()]
    Param(
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Site code")]
              $SiteCode,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Name")]
              $SiteServer,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Folder Name")]
              $Name,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Folder Object Type")]
              $ObjectType,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter parent folder ID")]
              $ParentContainerNodeId                                          
         )
 
    $Arguments = @{Name = $Name; ObjectType = "$ObjectType"; ParentContainerNodeId = "$ParentContainerNodeId"}
    Try{
        Set-WmiInstance -Namespace "root\sms\Site_$SiteCode" -Class "SMS_ObjectContainerNode" -Arguments $Arguments `
        -ComputerName $SiteServer -ErrorAction STOP
    }
    Catch{
        $date = get-date     
            $_.Exception.Message
        $date , "IMPORT MODULE FAILED - New Folder" | out-file  C:\Temp\IMPORT.log -append
    }          
}

#---------------------------------------------------Get-ParentNode-------------------------------------------------------------------------------------
 
#Function for Fetching the node
#This function is used to fetch the node of folder under which all the Vendor Folders would be created.In this case "SRT-Packages"
function Get-ParentNode($ConsoleFolder)
{
 
$Confolder=$ConsoleFolder.split("\")
$ParentContainerNodeId="0"
$len=$Confolder.length 
 
 
For ($i=0; $i -lt $len; $i++)
{
 
New-CMFolder -SiteCode $SiteCode -SiteServer $SiteServer -Name $Confolder[$i] -ObjectType $ObjectType -ParentContainerNodeId $ParentContainerNodeId
$Parent_Node1=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $Confolder[$i]) -and ($_.ObjectType -like '6000') }
$ParentContainerNodeId = $Parent_Node1.ContainerNodeID
}
$len=$len-1
$Parent_Node=Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms\site_$siteCode" -Query "Select * From SMS_ObjectContainerNode" | Where { ($_.Name -like $Confolder[$len]) -and ($_.ObjectType -like '6000') }
$ParentContainerNodeId1 = $Parent_Node.ContainerNodeID
 
$obj= new-object psobject
 
$obj | Add-Member Noteproperty ParentContainerNodeId $ParentContainerNodeId1
Return $obj
}
#----------------------------------------------------Primary User------------------------------------------------------------

Function Get-PrimaryUser
{
$oSettingReference = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.GlobalSettingReference  -ArgumentList ( "GLOBAL",
"PrimaryDevice",
[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType]::Boolean,
"PrimaryDevice_Setting_LogicalName",
[Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ConfigurationItemSettingSourceType]::CIM)

$oAnnotation             = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Annotation      
$oAnnotation.DisplayName = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.LocalizableString -ArgumentList "DisplayName", "Primary device Equals True", $null

$oConstantValue = New-Object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.ConstantValue([System.Convert]::ToBoolean($True),
[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.DataType]::Boolean)

$operands = new-object "Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection``1[[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.ExpressionBase]]"
 
$operands.Add($oSettingReference)
$operands.Add($oConstantValue)

$oExpression= new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.Expression([Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::IsEquals, $operands)

$PrimaryUser = new-object "Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Rule" -ArgumentList `
            ("Rule_" + [Guid]::NewGuid().ToString()), 
            ([Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::None), $oAnnotation,$oExpression

return $PrimaryUser
}

#----------------------------------------------------OSRule------------------------------------------------------------
Function Get-OSRule
{
$oOperands = new-object "Microsoft.ConfigurationManagement.DesiredConfigurationManagement.CustomCollection``1[[Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.RuleExpression]]"
$oOperands.Add("Windows/All_x64_Windows_10_and_higher_Clients")
$oOperands.Add("Windows/All_ARM_Windows_8.1_Client")
$oOperands.Add("Windows/All_x64_Windows_8.1_Client")

$oOSExpression           = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Expressions.OperatingSystemExpression -ArgumentList ([Microsoft.ConfigurationManagement.DesiredConfigurationManagement.ExpressionOperators.ExpressionOperator]::OneOf), $oOperands    
 
$oAnnotation             = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Annotation      
$oAnnotation.DisplayName = new-object Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.LocalizableString -ArgumentList "DisplayName", "Primary device Equals True", $null


$oDTRule = new-object "Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.Rule" -ArgumentList `
            ("Rule_" + [Guid]::NewGuid().ToString()), 
            ([Microsoft.SystemsManagementServer.DesiredConfigurationManagement.Rules.NoncomplianceSeverity]::None), $oAnnotation,$oOSExpression

return $oDTRule
}
#----------------------------------------------------Site-Code------------------------------------------------------------

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

#----------------------------------------------------Move-CMObject------------------------------------------------------------
 
#Function to Move Object.
#When you import the application it would be under root directory.Move function is used to move to corresponding Vendor folder.
 
 
Function Move-CMObject
{
    [CmdLetBinding()]
    Param(
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Site code")]
              $SiteCode,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Site Server Name")]
              $SiteServer,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter Object ID")]
              [ARRAY]$ObjectID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter current folder ID")]
              [uint32]$CurrentFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Enter target folder ID")]
              [uint32]$TargetFolderID,
    [Parameter(Mandatory=$True,HelpMessage="Please Ente .r obje .ct type ID")]
              [uint32]$ObjectTypeID              
        )
 
    
        Invoke-WmiMethod -Namespace "Root\SMS\Site_$SiteCode" -Class SMS_objectContainerItem -Name MoveMembers -ArgumentList $CurrentFolderID,$ObjectID,$ObjectTypeID,$TargetFolderID -ComputerName $SiteServer -ErrorAction STOP
     
 
}
#-------------------------------Script-------------------------------------------------------------------------------------------------------------------



#Global Variable Declaration

Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.ApplicationManagement.dll"
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.ApplicationManagement.MsiInstaller.dll"
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.ManagementProvider.dll"
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.ApplicationManagement.Extender.dll"
#used for creating rules
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\DcmObjectModel.dll"
#WQL Connection to Server
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\AdminUI.WqlQueryEngine.dll"
#Application Wrapper and Factory
Add-Type -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\AdminUI.AppManFoundation.dll"

$SiteCode = Get-SiteCode($SiteServer)
$ObjectType="6000"
$ConsoleFolder="APP Packages"
$App_Source_Location ="\\vnwlscmpm01\contentsourceg$\Applications\" 
$csv = import-csv -Path $CSV_FILE
$csv | ForEach {
$x = [String]$_.ApplicationName.Trim(" ")
$AppVFilePath = $App_Source_Location + $x + "\Package"
$oDTRule = Get-OSRule
$PrimaryUser = Get-PrimaryUser

#SCCM Powershell Window
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
           $date = get-date     
            $_.Exception.Message
        $date , "IMPORT MODULE FAILED" + $Vendor| out-file  C:\Temp\IMPORT.log -append
            }

$Setsite = $SiteCode + ":"
CD $Setsite

#Console folder parent node fetch

 If($ConsoleFolder.Length -gt 0) # Valid
    {
       $Confolders=$ConsoleFolder.split("\")
$len=$Confolders.length 
$Parent_Node_obj=Get-ParentNode($ConsoleFolder)

$Parent_Nodes=$Parent_Node_obj.ParentContainerNodeID
$Parent_Node =$Parent_Nodes[$len]
If($Parent_Node -eq $null){

$Parent_Node=$Parent_Node_obj.ParentContainerNodeID
}
    }
    Else # Invalid
    {
             $Parent_Node = 0
    }
#New Application creation on the console
New-CMApplication -Name $_.ApplicationName -Owner "Matha" -SupportContact "Matha" -LocalizedApplicationName $_.ApplicationDisplayName -Description "AppV" -AutoInstall $true
Set-CMApplication -Name $_.ApplicationName -AppCategory "App-V" 
Set-Location C:\

$AppVFile = Get-ChildItem $AppVFilePath -Recurse | where {$_.extension -eq ".appv"}

If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
$Setsite = $SiteCode + ":"
CD $Setsite

Add-CMDeploymentType -ApplicationName $_.ApplicationName -AppV5xInstaller -AutoIdentifyFromInstallationFile -ForceForUnknownPublisher $true -InstallationFileLocation $AppVFile.fullname  -DeploymentTypeName $_.ApplicationName
$appName = $_.ApplicationName.trim(" ")

Set-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $appName -AddRequirement $oDTRule
Set-CMDeploymentType -ApplicationName $appName -DeploymentTypeName $appName -AddRequirement $PrimaryUser
Start-CMContentDistribution -ApplicationName $_.ApplicationName -DistributionPointGroupName  "All NW DPs - Software Distribution"

$TargetFolderID = Get-WmiObject -Namespace "root\sms\Site_$SiteCode" -ComputerName $SiteServer -Class SMS_ObjectContainerNode -Filter "Name = 'APP Packages'" | where {$_.objecttype -eq 6000}
$APPScopeID = Get-WmiObject -Namespace "root\sms\Site_$SiteCode" -ComputerName $SiteServer -Class SMS_ApplicationLatest -Filter "LocalizedDisplayName='$appName'"
Move-CMObject -SiteCode $SiteCode -SiteServer $SiteServer -ObjectID $APPScopeID.ModelName -CurrentFolderID 0 -TargetFolderID $TargetFolderID.ContainerNodeID -ObjectTypeID $ObjectType
    }
 