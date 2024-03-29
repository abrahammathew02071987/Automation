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

#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "SCCM 2012 APPLICATION IMPORT                by Abraham Mathew"
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
      If($objTextBox.Text.Length -gt 0) # Valid
    {
        $VendorFolderLocation=$objTextBox.Text
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
$objLabel3.Text = "Import File location"
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

$VendorFolderLocation=$objTextBox.Text
$ConsoleFolder=$objTextBox1.Text
$SiteServer=$objTextBox2.Text

    If($VendorFolderLocation -eq "" -or $SiteServer -eq "" ) # Valid
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
 
    
        Invoke-WmiMethod -Namespace "Root\SMS\Site_$SiteCode" -Class SMS_objectContainerItem -Name MoveMembers -ArgumentList $CurrentFolderID,$ObjectID,$ObjectTypeID,$TargetFolderID -ComputerName $SiteServer -ErrorAction STOP
     
 
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
        $_.Exception.Message
        "THE VENDOR FOLDER PRESENT : " + $Name | out-file  C:\Windows\Logs\IMPORT.log -append
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
#-------------------------------Script-------------------------------------------------------------------------------------------------------------------



#Global Variable Declaration
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
        "IMPORT MODULE FAILED" + $Vendor| out-file  C:\Windows\Logs\IMPORT.log -append
            }

$Setsite = $SiteCode + ":"
CD $Setsite
$ObjectType="6000"
#$ConsoleFolder="Global Package\SRT-Packages"
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
 
Set-Location C:\
$a= Get-Childitem  -path $VendorFolderLocation
If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
$Setsite = $SiteCode + ":"
CD $Setsite

 
 
 
foreach($x in $a)  {
$APP_VAR =[String]$x.Name
$Vendor = $APP_VAR.trim(" ")


#---------------------------------------------------Vendor Folder Creation Under Console-------------------------------------------------
try 
            {
New-CMFolder -SiteCode $SiteCode -SiteServer $SiteServer -Name $Vendor -ObjectType $ObjectType -ParentContainerNodeId $Parent_Node
"Vendor Folder Creation: " + $X | out-file  C:\Windows\Logs\IMPORT.log -append
}
        catch 
            {      
            $_.Exception.Message
        "Folder Creation" + $X | out-file  C:\Windows\Logs\IMPORT.log -append
            }
$ExportFilesLocation=$VendorFolderLocation + $x.Name + "\*.*"
Set-Location C:\
$b =Get-Childitem  -path $ExportFilesLocation  -include *.zip 
If(Test-Path "c:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin"
}
If(Test-Path "c:\Program Files\Microsoft Configuration Manager\AdminConsole\bin"){
cd "Program Files\Microsoft Configuration Manager\AdminConsole\bin"
}
$Setsite = $SiteCode + ":"
CD $Setsite


foreach($y in $b)  {
$APP_VAR1 =[String]$y.Name
$APP_VAR2 = $APP_VAR1.trim(" ")
$ApplicationName = $APP_VAR2.TrimEnd(".zip")
"Application Name: " + $y | out-file  C:\Windows\Logs\IMPORT.log -append
$Applications = Get-WmiObject -Namespace "root\sms\Site_$SiteCode" -ComputerName $SiteServer -Class SMS_ApplicationLatest
If($Applications -eq $null ) {
#Import Application to the SCCM Console
try 
            {
                Import-CMApplication -FilePath $y.FullName -ImportActionType DIRECTiMPORT
                "IMPORT sucessful: " + $y | out-file  C:\Windows\Logs\IMPORT.log -append
    }
        catch 
            {      
            $_.Exception.Message
        "IMPORT FAILURE" + $y.FullName | out-file  C:\Windows\Logs\IMPORT.log -append
         
            }
}
else {
If($Applications.LocalizedDisplayName -eq $ApplicationName)
{
#Import Application to the SCCM Console
try 
            {
                Import-CMApplication -FilePath $y.FullName -ImportActionType OVERWRITE 
                "IMPORT sucessful: " + $y | out-file  C:\Windows\Logs\IMPORT.log -append
    }
        catch 
            {      
            $_.Exception.Message
        "IMPORT FAILURE" + $y.FullName | out-file  C:\Windows\Logs\IMPORT.log -append
            }
}
else {

#Import Application to the SCCM Console
try 
            {
                Import-CMApplication -FilePath $y.FullName -ImportActionType DIRECTiMPORT
                "IMPORT sucessful: " + $y | out-file  C:\Windows\Logs\IMPORT.log -append
    }
        catch 
            {      
            $_.Exception.Message
        "IMPORT FAILURE" + $y.FullName | out-file  C:\Windows\Logs\IMPORT.log -append
try 
            {
         Import-CMApplication -FilePath $y.FullName
 }
        catch 
            {      
            $_.Exception.Message
        "IMPORT FAILURE" + $y.FullName | out-file  C:\Windows\Logs\IMPORT.log -append
}
 
}            
}
}
$APPScopeID = Get-WmiObject -Namespace "root\sms\Site_$SiteCode" -ComputerName $SiteServer -Class SMS_ApplicationLatest -Filter "LocalizedDisplayName='$ApplicationName'"           
$TargetFolderID = Get-WmiObject -Namespace "root\sms\Site_$SiteCode" -ComputerName $SiteServer -Class SMS_ObjectContainerNode -Filter "Name='$Vendor'"
 
try 
            { 

Move-CMObject -SiteCode $SiteCode -SiteServer $SiteServer -ObjectID $APPScopeID.ModelName -CurrentFolderID 0 -TargetFolderID $TargetFolderID.ContainerNodeID -ObjectTypeID $ObjectType
"Application moved successfully:" + $y | out-file  C:\Windows\Logs\IMPORT.log -append
}
        catch 
            {      
            $_.Exception.Message
        "APPLICATION WAS NOT MOVED SUCCESSFULLY :" + $y.FullName | out-file  C:\Windows\Logs\IMPORT.log -append
            }    
}

}
EXIT(0)