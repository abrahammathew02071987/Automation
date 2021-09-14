
Param
(
    [parameter(Mandatory=$true)]
    [String]$ServerInstance,
    [parameter(Mandatory=$true)]
    [String]$Database,
    [parameter(Mandatory=$true)]
    [String]$Tablename,
    [parameter(Mandatory=$false)]
    [String]$Schema = "dbo" ,
    [String]$Username = "NextGenITBui_so" ,
    [String]$pwd = "2DyMcVb14RcHdG7"
)




#Install-Module -Name SqlServer -AllowClobber
If(-not(Get-InstalledModule SQLServer  -ErrorAction silentlycontinue)){
    Get-PackageProvider -Name "nuget" -ForceBootstrap
    Install-Module SQLServer -Confirm:$False -Force -AllowClobber
}
Import-Module SqlServer
$password = ConvertTo-SecureString $pwd -AsPlainText -Force
$cred =  New-Object System.Management.Automation.PSCredential ($Username,$password )
$CSVImport = Import-Csv "C:\EnterpriseTestbed\Result\CHM_WMI.csv"

ForEach ($CSVLine in $CSVImport)
{
# Setting variables for the CSV line, ADD ALL 170 possible CSV columns here
$ID = $CSVLine.ID
$ComputerName = $CSVLine.ComputerName
$Model_Name = $CSVLine.Model_Name
$xyz = $CSVLine.Platform
$Platform = $xyz.trimend()
$Platform_Generation = $CSVLine.'Platform Generation'
$CPU = $CSVLine.CPU
$RAM = $CSVLine.RAM
$PTLC = $CSVLine.'Processor Total Logical Cores'
$BIOSVersion = $CSVLine.'BIOS Version'
$SMBIOBIOS_Version = $CSVLine.SMBIOBIOS_Version
$OSShortversion = $CSVLine.OS

$ChromeVersion = $CSVLine.ChromeVersion
$Wireless_Driver_Version = $CSVLine.Wireless_Driver_Version
$Bluetooth_Driver_Version = $CSVLine.Bluetooth_Driver_Version
$Graphics_Driver_Version = $CSVLine.Graphics_Driver_Version
$RealTek_Audio_Driver_Version = $CSVLine.RealTek_Audio_Driver_Version



##############################################
# SQL INSERT of CSV Line/Row
##############################################
$SQLInsert = "USE $Database
INSERT INTO $Tablename (ComputerName,Model_Name,Platform,[Platform Generation],CPU,RAM,[Processor Total Logical Cores],[BIOS Version],SMBIOBIOS_Version,OS,ChromeVersion,Wireless_Driver_Version,Bluetooth_Driver_Version,Graphics_Driver_Version,RealTek_Audio_Driver_Version)
VALUES( '$ComputerName', '$Model_Name', '$Platform', '$Platform_Generation', '$CPU', '$RAM', '$PTLC', '$BIOSVersion', '$SMBIOBIOS_Version', '$OSShortversion', '$ChromeVersion', '$Wireless_Driver_Version', '$Bluetooth_Driver_Version', '$Graphics_Driver_Version', '$RealTek_Audio_Driver_Version');"
# Running the INSERT Query
Invoke-SQLCmd -Query $SQLInsert -ServerInstance $ServerInstance -Credential $cred
# End of ForEach CSV line below
}