function Test-PsRemoting
{
    param(
        [Parameter(Mandatory = $true)]
        $computername
    )
   
    try
    {
        $errorActionPreference = "Stop"
        $result = Invoke-Command -ComputerName $computername { 1 }
    }
    catch
    {
        Write-Verbose $_
        return $false
    }
   
    ## I've never seen this happen, but if you want to be
    ## thorough....
    if($result -ne 1)
    {
        Write-Verbose "Remoting to $computerName returned an unexpected result."
        return $false
    }
   
    $true   
}

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




function reg-write{
    param(
        [Parameter(Mandatory = $true)]
        $reg
    )

$R = Invoke-Command -Session $session -ScriptBlock{ 
$item = New-Object System.Object
$folderExist = test-path $($using:reg)
If($folderExist -eq $True){

[System.Collections.ArrayList]$sb = @()
$drive = New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR
$item | Add-Member -MemberType NoteProperty -Name "Root" -Value (Get-ItemProperty $($using:reg) | Select-Object * -ExcludeProperty PS*)
Get-ChildItem $($using:reg) -Recurse | ForEach-Object { 

$sb.Add( @(Get-ItemProperty $_.pspath | Select-Object * -exclude PS* ))

}
write-host $sb[1]  
$item | Add-Member -MemberType NoteProperty -Name "SubRoot" -Value $sb
$item
}else {
[bool]($null)
}
}
""   | format-list |Out-File $RegFilename
write-host $R
if($R -ne ""){
$R.Root   | format-list |Out-File $RegFilename -Append
for ( $index = 0; $index -lt $r.subroot.count; $index++ )    {
        $r.subroot[$index] | format-list | out-file $RegFilename -append
}
}else{

"Registery $reg Missing" | out-file $RegFilename -append
}

}









#---------------------------------------------------PowerShell GUI----------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
#$objForm | Get-Member -Force
Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::OK
$MessageboxTitle = “CLIENT STATUS”
$Messageboxbody = “Machine is Online.”
$Messageboxbody1 = “Machine is Offline.”
$MessageboxTitle1 = “FORM STATUS”
$Messageboxbody2 = “FORM SUBMITTED.”
$Messageboxbody3 = “COMPLETE!.”
$MessageIcon = [System.Windows.MessageBoxImage]::Error
$MessageIcon1 = [System.Windows.MessageBoxImage]::Information

$objForm = New-Object System.Windows.Forms.Form 
$objForm.AutoScale = $true
$resizeHandler = { "form resized" }

$objform.Add_Resize( $resizeHandler )
$objForm.Text = "Client Machine Configuration Reporting Tool                                                         by Abraham Mathew"
$objForm.Size = New-Object System.Drawing.Size(1100,650) 
$objForm.StartPosition = "CenterScreen"
[bool]$script:Quit = $false


$objForm.BackgroundImageLayout = "Zoom"
    # None, Tile, Center, Stretch, Zoom

$Font = New-Object System.Drawing.Font("Times New Roman",16,[System.Drawing.FontStyle]::Regular)
    # Font styles are: Regular, Bold, Italic, Underline, Strikeout
    
$objForm.Font = $Font


$OKButton = New-Object System.Windows.Forms.Button
$OKButton.AutoSize = $true
$OKButton.Location = New-Object System.Drawing.Size(300,500)
$OKButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$OKButton.Text = "Ok"
$OKButton.font = $LabelFont
#$OKButton.Add_Click({$VendorFolderLocation=$objTextBox.Text;$ConsoleFolder=$objTextBox1.Text;$SiteServer=$objTextBox2.Text;$objForm.Close()})
$OKButton.Add_Click({
   If($objTextBox.Text.Length -gt 0){
    $result=[System.Windows.MessageBox]::Show($Messageboxbody2,$MessageboxTitle1,$ButtonType,$messageicon1)
    }
#---------------------------------------------------
$Clientname=$objTextBox.Text
$Logpath=$objTextBox2.Text
$DestinationPath=$objTextBox4.Text
$RegistryPath=$objTextBox3.Text
$NewLogfileName = $Clientname + ".log"
$NewLogfile = "c:\temp\" + $NewLogfileName
$DestinationFolder = $DestinationPath.TrimEnd("\") + "\" + $Clientname
$logfolder = $DestinationFolder + "\" + "Logs"
$Registry = $DestinationFolder + "\" + "Registry"
if($DestinationPath){
$NewLogfile = $DestinationFolder + "\" + $NewLogfileName
}
    

Write-Host $Clientname
Write-Host $Logpath
Write-Host $DestinationPath
Write-Host $RegistryPath



If($Clientname -ne $null) {

$folderExist = test-path $DestinationFolder
If($folderExist -eq $false){
 New-Item -Path $DestinationFolder -ItemType Directory

}

 
$session = New-PSSession -ComputerName $Clientname
if($Logpath){
$folderExist = test-path $logfolder
If($folderExist -eq $false){

 New-Item -Path $logfolder -ItemType Directory
 
}
$a = Get-Content -Path $Logpath  | %{ 

$Log_instance = $_ 
$LogFilename = $logfolder + "\" + $Log_instance.Split("\")[-1]
 
$L = Invoke-Command -Session $session -ScriptBlock{
$folderExist = test-path $($using:Log_instance) 
If($folderExist -eq $True){
Get-Content -Path $($using:Log_instance) 
}else {
$item = New-Object System.Object
$item | Add-Member -MemberType NoteProperty -Name "Eroor" -Value "Logpath $($using:Log_instance) missing"
$item
}
}
$L | Out-File $LogFilename 
}


}
if($RegistryPath){
$folderExist = test-path $Registry
If($folderExist -eq $false){


 New-Item -Path $Registry -ItemType Directory
 }
$RegFilename = $Registry + "\" + "Registry.log"
if($radioButton1.Checked -eq "True"){

$reg =  "HKLM:" + "\" + $RegistryPath
reg-write($reg)
}else {
$reg =  "HKCR:" + "\" + $RegistryPath
reg-write($reg)
}


}
}
    $script:Quit = $false
    $result=[System.Windows.MessageBox]::Show($Messageboxbody3,$MessageboxTitle1,$ButtonType,$messageicon1)

    })
$objForm.Controls.Add($OKButton)

#---------------------------------------------------------------------------------------------------ok Button

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(420,500)
$CancelButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$CancelButton.Text = "Cancel"
$CancelButton.font = $LabelFont
$CancelButton.AutoSize = $true
$CancelButton.Add_Click({$objForm.Close(); $cancel = $true;$script:Quit = $True })
$objForm.Controls.Add($CancelButton)

#-----------------------------------------test button---------------------------------------------------------------

$TestButton = New-Object System.Windows.Forms.Button
$TestButton.AutoSize = $true
$TestButton.Location = New-Object System.Drawing.Size(920,100)
$TestButton.Size = New-Object System.Drawing.Size(110,30)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$TestButton.Text = "Test"
$TestButton.font = $LabelFont
#$TestButton.Add_Click({$VendorFolderLocation=$objTextBox.Text;$ConsoleFolder=$objTextBox1.Text;$SiteServer=$objTextBox2.Text;$objForm.Close()})
$TestButton.Add_Click({
    If($objTextBox.Text.Length -gt 0) # Valid
    {
        $ClientName=$objTextBox.Text
        $client_status = Test-PsRemoting $ClientName
        If($client_status)
        {
            $result=[System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon1)
        }else
        {
            $result=[System.Windows.MessageBox]::Show($Messageboxbody1,$MessageboxTitle,$ButtonType,$messageicon)
        }
    }
    Else # Invalid
    {
             [windows.forms.messagebox]::show($objLabel1.Text,"Enter Input")
    }
    
    
    })
$objForm.Controls.Add($TestButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.AutoSize = $true
$LabelFont = New-Object System.Drawing.Font("Times New Roman",20,[System.Drawing.FontStyle]::Bold)
$objLabel.Location = New-Object System.Drawing.Size(100,30)
$objLabel.Size = New-Object System.Drawing.Size(630,40) 
$objLabel.Text = "Please enter the input to the form"
$objLabel.font = $LabelFont


$objForm.Controls.Add($objLabel) 

$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.AutoSize = $true
$objTextBox.Location = New-Object System.Drawing.Size(400,100) 
$objTextBox.Size = New-Object System.Drawing.Size(500,20)
$objForm.Controls.Add($objTextBox)

#$objLabel = New-Object System.Windows.Forms.Label
#$objLabel.Location = New-Object System.Drawing.Size(300,205) 
#$objLabel.Size = New-Object System.Drawing.Size(400,40)
#$LabelFont = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular) 
#$objLabel.Text = "Registry Path"
#$objForm.Controls.Add($objLabel) 
#$objLabel.font = $LabelFont 

$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.AutoSize = $true
$objLabel3.Location = New-Object System.Drawing.Size(100,100) 
$objLabel3.Size = New-Object System.Drawing.Size(280,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel3.Text = "Computer Name"
$objLabel3.font = $LabelFont
$objForm.Controls.Add($objLabel3) 
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(100,250)
$objLabel.Size = New-Object System.Drawing.Size(280,40)
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular)
$objLabel.Text = "Destination Dir"
$objLabel.font = $LabelFont
$objForm.Controls.Add($objLabel)




$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.AutoSize = $true
$objTextBox2.Location = New-Object System.Drawing.Size(400,175) 
$objTextBox2.Size = New-Object System.Drawing.Size(500,20) 
$objForm.Controls.Add($objTextBox2)



$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.AutoSize = $true
$objLabel2.Location = New-Object System.Drawing.Size(100,175) 
$objLabel2.Size = New-Object System.Drawing.Size(250,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel2.Text = "Log FilePath"
$objForm.Controls.Add($objLabel2)
$objLabel2.font = $LabelFont



$objTextBox4 = New-Object System.Windows.Forms.TextBox 
$objTextBox4.AutoSize = $true
$objTextBox4.Location = New-Object System.Drawing.Size(400,250) 
$objTextBox4.Size = New-Object System.Drawing.Size(500,20) 
$objForm.Controls.Add($objTextBox4)

#--------------------------Registry Path--------------------------------------------------------------

$objLabel5 = New-Object System.Windows.Forms.Label
$objLabel5.AutoSize = $true
$objLabel5.Location = New-Object System.Drawing.Size(100,325) 
$objLabel5.Size = New-Object System.Drawing.Size(250,40) 
$LabelFont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::Regular) 
$objLabel5.Text = "Registry Path"
$objForm.Controls.Add($objLabel5)
$objLabel5.font = $LabelFont

$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.AutoSize = $true
$objTextBox3.Location = New-Object System.Drawing.Size(400,400) 
$objTextBox3.Size = New-Object System.Drawing.Size(500,20) 
$objForm.Controls.Add($objTextBox3)

$radioButton1 = New-Object system.windows.Forms.RadioButton 
$radioButton1.Text = "HKLM"
$radioButton1.AutoSize = $true
$radioButton1.Width = 104
$radioButton1.Height = 20
$radioButton1.location = new-object system.drawing.point(400,325)
$radioButton1.Font = "Microsoft Sans Serif,10"
$objForm.controls.Add($radioButton1)

$radioButton2 = New-Object system.windows.Forms.RadioButton
$radioButton2.Text = "HKCR"
$radioButton2.AutoSize = $true
$radioButton2.Width = 104
$radioButton2.Height = 20
$radioButton2.location = new-object system.drawing.point(500,325)
$radioButton2.Font = "Microsoft Sans Serif,10"
$objForm.controls.Add($radioButton2)
$objForm.ControlBox = $False;

$objForm.Add_Shown({$objForm.Activate()})

[void] $objForm.ShowDialog()


if($script:Quit -eq 'True'){
exit(0)
}







