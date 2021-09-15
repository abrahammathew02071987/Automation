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

if ($env:Processor_Architecture -ne "x86")  

{ write-warning 'Launching x86 PowerShell'

&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noninteractive -noprofile -file $myinvocation.Mycommand.path -executionpolicy bypass

exit

}

"Always running in 32bit PowerShell at this point."

$env:Processor_Architecture

[IntPtr]::Size

Import-Module AppVPkgConverter

function GetBrowseLocation

{

    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $browse = New-Object System.Windows.Forms.FolderBrowserDialog

    $browse.RootFolder = [System.Environment+SpecialFolder]'MyComputer'

    $browse.ShowNewFolderButton = $false

    $browse.Description = "Choose a directory"

 

  

    

        if ($browse.ShowDialog() -eq "OK")

        {

            $loop = $false

        }

 

    $browse.SelectedPath

    $browse.Dispose()

}

 

function filebrwse

{

 

[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    [System.Windows.Forms.Application]::EnableVisualStyles()

 

$fd = New-Object System.Windows.Forms.OpenFileDialog

$fd.InitialDirectory="c:\"

$fd.CheckFileExists=$true

 

$fd.Multiselect=$false

  if ($fd.ShowDialog() -eq "OK")

        {

            $fd.filename

        }

}

 

 

function cnvrt

{

 

param($srcpath,$dstpath,$xmlpath)

 

write-host $srcpath

write-host $dstpath

write-host $xmlpath

 

if (!(test-path -Path "$dstpath" -Type Container))

{

mkdir "$dstpath"

}

 

$xl=new-object -ComObject "Excel.Application"

$xl.visible=$false

$wb = $xl.Workbooks.Open("$xmlpath")

$ws = $wb.Sheets.Item("Test-Results")

 

$xlbooks =$xl.workbooks.Add()

$sht1=$xlbooks.worksheets.item(1)

$xlbooks =$xl.workbooks.Add()

$sht1=$xlbooks.worksheets.item(1)

$sht1.Name="Test-Results"

$cls = $sht1.Cells

 

$cls.item(1,1)="Package Name"

$cls.item(1,2)="Converted"

$cls.item(1,3)="Errors"

$cls.item(1,4)="Warnings"

$cls.item(1,5)="Information"

$cls.item(1,1).font.bold=$True

$cls.item(1,1).font.size=12

$cls.item(1,2).font.bold=$True

$cls.item(1,2).font.size=12

$cls.item(1,3).font.bold=$True

$cls.item(1,3).font.size=12

$cls.item(1,4).font.bold=$True

$cls.item(1,4).font.size=12

$cls.item(1,5).font.bold=$True

$cls.item(1,5).font.size=12

 

$rn=$ws.UsedRange.Rows.count

$rw=1

$cl=1

 

 

for ($i = 2; $i -le $rn; $i++) {

 

 

  if ( $ws.Cells.Item($i, 2).text -eq "PASS" ) {

 

  $rw+=1

 

  $cls.item($rw,1)=$ws.Cells.Item($i, 1).text

  $pname=$ws.Cells.Item($i, 1).text

 

 

if (!(test-path -Path "$dstpath\$pname" -Type Container))

{

mkdir "$dstpath\$pname"

}

 


 $crs=ConvertFrom-AppvLegacyPackage -SourcePath "$srcpath\$pname" -destinationpath "$dstpath\$pname" -ErrorAction Ignore

 

if(($crs.Errors.count -eq 0) -and ($crs.Warnings.count -eq 0))

{

$cls.item($rw,2)="PASSED"

 

}

else

{

$cls.item($rw,2)="FAILED"

$cerval=""

$cercnt=$crs.Errors.count

for($j=0;($j -lt $cercnt) -and ($crs.Errors.count -ne 0) ;$j++){ $cerval+=$crs.Errors.item($j).tostring() }

$cwarval=""

$cwarcnt=$crs.Warnings.count

for($j=0;($j -lt $cwarcnt)  -and ($crs.Warnings.count -ne 0) ;$j++){ $cwarval+=$crs.Warnings.item($j).tostring() }

 

$cinfval=""

$cinfcnt=$crs.Information.count

for($j=0;($j -lt $cinfcnt) -and ($crs.Information.count -ne 0) ;$j++){ $cinfval+=$crs.Information.item($j).tostring() }

 

 

 

$cls.item($rw,3)=$cerval

$cls.item($rw,4)=$cwarval

$cls.item($rw,5)=$cinfval

}

 

 

  }

}

 

$xlbooks.saveas("c:\convert.xlsx")

[System.Windows.Forms.MessageBox]::Show("Report saved as c:\convert.xlsx")

$xlbooks.close()

$wb.Close()

$xl.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)

}

function test

{

param($fldrpath)

 

 

 

if(test-path -Path "$fldrpath" -PathType Container)

{

 

$xl=new-object -ComObject "Excel.Application"

$xl.visible=$false

$xlbooks =$xl.workbooks.Add()

$sht1=$xlbooks.worksheets.item(1)

$sht1.Name="Test-Results"

$cls = $sht1.Cells

 

$cls.item(1,1)="Package Name"

$cls.item(1,2)="Is Convertable"

$cls.item(1,3)="Erros"

$cls.item(1,4)="warnings"

$cls.item(1,5)="Information"

 

$cls.item(1,1).font.bold=$True

$cls.item(1,1).font.size=12

$cls.item(1,2).font.bold=$True

$cls.item(1,2).font.size=12

$cls.item(1,3).font.bold=$True

$cls.item(1,3).font.size=12

$cls.item(1,4).font.bold=$True

$cls.item(1,4).font.size=12

$cls.item(1,5).font.bold=$True

$cls.item(1,5).font.size=12

$rw=2

$cl=1

 

$subfldr=Get-ChildItem -Path "$fldrpath"

 

foreach ($pkg in $subfldr)

{

 

$name=$pkg.FullName

 

$rs=Test-AppvLegacyPackage -SourcePath "$name"

 

$cls.item($rw,1)="$pkg"

 

if(($rs.Errors.count -eq 0) -and ($rs.Warnings.count -eq 0))

{

$cls.item($rw,2)="PASS"

 

}

else

{

$cls.item($rw,2)="FAIL"

$erval=""

$ercnt=$rs.Errors.count

for($i=0;($i -lt $ercnt) -and ($rs.Errors.count -ne 0) ;$i++){ $erval+=$rs.Errors.item($i).tostring() }

$warval=""

$warcnt=$rs.Warnings.count

for($i=0;($i -lt $warcnt)  -and ($rs.Warnings.count -ne 0) ;$i++){ $warval+=$rs.Warnings.item($i).tostring() }

 

$infval=""

$infcnt=$rs.Information.count

for($i=0;($i -lt $infcnt) -and ($rs.Information.count -ne 0) ;$i++){ $infval+=$rs.Information.item($i).tostring() }

 

 

 

$cls.item($rw,3)=$erval

$cls.item($rw,4)=$warval

$cls.item($rw,5)=$infval

}

 

$rw+=1

 

}

 

$xlbooks.saveas("c:\test.xlsx")

$xlbooks.close()

$xl.quit()

 

[System.Windows.Forms.MessageBox]::Show("Report saved as c:\test.xlsx")

 

}

 

else

 

{

 

[System.Windows.Forms.MessageBox]::Show($tbox1.text+"  folder does not exist")

 

}

}

 

$handler_slcttstfldr_Click=

{

 

    $fldrpath=GetBrowseLocation

 

    $tbox1.text=$fldrpath

}

 

 

$handler_tstfldr_Click=

 

{

 

test($tbox1.text)

 

}

 

$handler_slctcnvrtfldr_Click=

{

 

    $fldrpath=GetBrowseLocation

 

    $tbox2.text=$fldrpath

}

 

 

$handler_cnvrtfldr_Click=

 

{

 

Cnvrt $tbox1.text $tbox2.text "c:\test.xlsx"

 

}

 

 

function GenerateForm {

 

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

 

$form1 = New-Object System.Windows.Forms.Form

$lbl1=New-Object System.Windows.Forms.Label

$lbl1.Text = "Click Browse to select  folder containing App-v 4.5 or above Packages"

$Slcttestfolder = New-Object System.Windows.Forms.Button

$tbox1=New-Object System.Windows.Forms.TextBox

$lbl2=New-Object System.Windows.Forms.Label

$tstfldr = New-Object System.Windows.Forms.Button

$lbl2.Text = "Click Browse to select  folder for saving Converted Packages"

$tbox2=New-Object System.Windows.Forms.TextBox

$slctcnvrtfldr = New-Object System.Windows.Forms.Button

$cnvrtfldr = New-Object System.Windows.Forms.Button

 

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

 

 

$OnLoadForm_StateCorrection=

{

    $form1.WindowState = $InitialFormWindowState

}

$form1.Text = "Conersion from App-v 4.5 or above to App-v 5.0 by Abraham Mathew"

$form1.Name = "form1"

$form1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Size = New-Object System.Drawing.Size

$System_Drawing_Size.Width = 500

$System_Drawing_Size.Height = 300

$form1.ClientSize = $System_Drawing_Size

 

 

$lbl1.Enabled=$true

$lbl1.Height=30

$lbl1.Width=400

$lbl1.TabIndex=1

$lbl1.Visible=$true

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 20

$System_Drawing_Point.Y = 25

$lbl1.Location = $System_Drawing_Point

$form1.Controls.Add($lbl1)

 

 

$Slcttestfolder.Enabled=$true

$Slcttestfolder.TabIndex = 2

$Slcttestfolder.Name = "slcttstfldr"

$System_Drawing_Size = New-Object System.Drawing.Size

$System_Drawing_Size.Width = 70

$System_Drawing_Size.Height = 30

$Slcttestfolder.Size = $System_Drawing_Size

$Slcttestfolder.UseVisualStyleBackColor = $True

 

$Slcttestfolder.Text = "Browse"

 

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 30

$System_Drawing_Point.Y = 70

$Slcttestfolder.Location = $System_Drawing_Point

$Slcttestfolder.DataBindings.DefaultDataSourceUpdateMode = 0

$Slcttestfolder.add_Click($handler_slcttstfldr_Click)

 

$form1.Controls.Add($Slcttestfolder)

 

$tbox1.Enabled=$true

$tbox1.Height=30

$tbox1.Width=200

$tbox1.text=""

$tbox1.TabIndex=3

$tbox1.Visible=$true

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 120

$System_Drawing_Point.Y = 70

$tbox1.Location = $System_Drawing_Point

$tbox1.Name = "textBox1"

 

$form1.Controls.Add($tbox1)

 

$tstfldr.Enabled=$true

$tstfldr.TabIndex = 4

$tstfldr.Name = "tstfldr"

$System_Drawing_Size = New-Object System.Drawing.Size

$System_Drawing_Size.Width = 70

$System_Drawing_Size.Height = 30

$tstfldr.Size = $System_Drawing_Size

$tstfldr.UseVisualStyleBackColor = $True

$tstfldr.Text = "Test"

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 360

$System_Drawing_Point.Y = 70

$tstfldr.Location = $System_Drawing_Point

$tstfldr.DataBindings.DefaultDataSourceUpdateMode = 0

$tstfldr.add_Click($handler_tstfldr_Click)

 

$form1.Controls.Add($tstfldr)

 

 

 

 

 

 

 

$lbl2.Enabled=$true

$lbl2.Height=30

$lbl2.Width=400

$lbl2.TabIndex=5

$lbl2.Visible=$true

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 20

$System_Drawing_Point.Y = 110

$lbl2.Location = $System_Drawing_Point

$form1.Controls.Add($lbl2)

 

 

$slctcnvrtfldr.Enabled=$true

$slctcnvrtfldr.TabIndex = 6

$slctcnvrtfldr.Name = "slctcnvrtfldr"

$System_Drawing_Size = New-Object System.Drawing.Size

$System_Drawing_Size.Width = 70

$System_Drawing_Size.Height = 30

$slctcnvrtfldr.Size = $System_Drawing_Size

$slctcnvrtfldr.UseVisualStyleBackColor = $True

 

$slctcnvrtfldr.Text = "Browse"

 

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 30

$System_Drawing_Point.Y = 145

$slctcnvrtfldr.Location = $System_Drawing_Point

$slctcnvrtfldr.DataBindings.DefaultDataSourceUpdateMode = 0

$slctcnvrtfldr.add_Click($handler_slctcnvrtfldr_Click)

 

$form1.Controls.Add($slctcnvrtfldr)

 

$tbox2.Enabled=$true

$tbox2.Height=30

$tbox2.Width=200

$tbox2.text=""

$tbox2.TabIndex=7

$tbox2.Visible=$true

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 120

$System_Drawing_Point.Y = 145

$tbox2.Location = $System_Drawing_Point

$tbox2.Name = "textBox2"

$tbox2.text="C:\Install\Appv 5.0"

 

$form1.Controls.Add($tbox2)

 

$cnvrtfldr.Enabled=$true

$cnvrtfldr.TabIndex = 8

$cnvrtfldr.Name = "cnvrtfldr"

$System_Drawing_Size = New-Object System.Drawing.Size

$System_Drawing_Size.Width = 70

$System_Drawing_Size.Height = 30

$cnvrtfldr.Size = $System_Drawing_Size

$cnvrtfldr.UseVisualStyleBackColor = $True

$cnvrtfldr.Text = "Convert"

$System_Drawing_Point = New-Object System.Drawing.Point

$System_Drawing_Point.X = 360

$System_Drawing_Point.Y = 145

$cnvrtfldr.Location = $System_Drawing_Point

$cnvrtfldr.DataBindings.DefaultDataSourceUpdateMode = 0

$cnvrtfldr.add_Click($handler_cnvrtfldr_Click)

 

$form1.Controls.Add($cnvrtfldr)

 

 

$InitialFormWindowState = $form1.WindowState

$form1.add_Load($OnLoadForm_StateCorrection)

$form1.ShowDialog()| Out-Null

 

}

 

 

GenerateForm