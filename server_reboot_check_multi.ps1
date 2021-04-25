Write-Host "Script Started at $([DateTime]::Now)." -BackgroundColor Black -ForegroundColor Cyan

##Variable Declaration
$server_loc="C:\temp\server.txt"
$server=gc $server_loc
$Report_path="C:\Scripts\Powershell\Server_reboot_checks"
$filename=$Report_path+"\Servers_Reboot_Checks.html"


##Removal Of Old HTML Report
if(Test-Path "$Report_path\*.html")
{
Remove-Item "$Report_path\*.html" -Force
}

$WINpassword = Get-Content "<path>\wintel_key.txt" | ConvertTo-SecureString -Key (Get-Content <path>\aes.key)
$WINcredential = New-Object System.Management.Automation.PsCredential("username",$WINPassword)

##HTML Table Formation
Write-Output '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">' | Out-File $filename
Write-Output '<html xmlns="http://www.w3.org/1999/xhtml">' | Out-File -Append $filename
Write-Output '<head>' | Out-File -Append $filename
Write-Output '<style>' | Out-File -Append $filename
Write-Output 'BODY{font-family: Arial; font-size: 10pt;}' | Out-File -Append $filename
Write-Output 'TABLE{border: 1px solid black; border-collapse: collapse;table-layout: auto;style="width: 100%;"}' | Out-File -Append $filename
Write-Output 'TH{border: 1px solid black; background: #86b6e3; padding: 5px;text-align:left;width:1%;white-space:nowrap;}' |Out-File -Append $filename
Write-Output 'TD{border: 1px solid black; padding: 5px;text-align:left;width:1%;white-space:nowrap;}' |Out-File -Append $filename
Write-Output '</style>' |Out-File -Append $filename
Write-Output '</head><body>' |Out-File -Append $filename
Write-Output '<table>'|Out-File -Append $filename
Write-Output '<colgroup><col/><col/></colgroup>'|Out-File -Append $filename
Write-Output "<h2> <u> Reboot Checks Details: </u> </h2>" |Out-File -Append $filename
Write-Output '<tr> <th>ServerName</th><th>Reboot Status</th><th>Uptime</th><th>Auto Services in Stopped State</th><th>C:\ Free space in %</th><th>Memory Usage(%)</th><th>CPUUsage(%)</th></tr>'|Out-File -Append $filename


##Reboot all servers at a time
#(Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -Credential $WINcredential).Win32Shutdown(6)
Restart-Computer -ComputerName $server -Credential $WINcredential -Verbose -Force
sleep 100


foreach ($vmname in (gc $server_loc)) 
{
try {
if (Test-Connection -ComputerName $vmname -Count 1 -ErrorAction SilentlyContinue)
{
$RebootStatus="Completed"
}
else
{
$RebootStatus="Failed"
}

##AutoService Stopped State
$autoservice=(Get-WmiObject -Class Win32_Service -ComputerName $vmname -Credential $WINcredential -ErrorAction Stop | Select-Object DisplayName,State,StartMode | Where-Object {$_.State -eq "Stopped" -and $_.StartMode -eq "Auto"}).DisplayName
$autoservice=$autoservice -join '<br/>'

##Uptime
$OS = Get-WmiObject win32_operatingsystem -ComputerName $vmname -Credential $WINcredential -ErrorAction Stop
$Uptime = $OS.ConvertToDateTime($os.LastBootUpTime)

##C Drive Free Percent Usage
$drivedetails=(Get-WmiObject Win32_LogicalDisk -ComputerName $vmname -Credential $WINcredential -ErrorAction Stop -Filter "DeviceID='C:'" | Select-Object SystemName,
@{ Name = "Drive" ; Expression = { ( $_.DeviceID ) } },
@{ Name = "SizeGB" ; Expression = {"{0:N1}" -f ( $_.Size / 1gb)}},
@{ Name = "FreeSpaceGB" ; Expression = {"{0:N1}" -f ( $_.Freespace / 1gb ) } },
@{ Name = "PercentFree" ; Expression = {"{0:P1}" -f ( $_.FreeSpace / $_.Size ) } })
$drive=$drivedetails.Drive
$total=$drivedetails.SizeGB
$Free=$drivedetails.FreeSpaceGB
$Free_perc=$drivedetails.PercentFree -replace '[%]'

##CPU Usage
$CPUavgLoad=(Get-WmiObject Win32_Processor -ComputerName $VMname -Credential $WINcredential -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average | Select Average).Average

##Memory Usage
$ComputerMemory=Get-WmiObject -Class win32_operatingsystem -ComputerName $vmname -Credential $WINcredential -ErrorAction Stop
$Total_physical_memory=[Math]::Round(($ComputerMemory.TotalVisibleMemorySize /1MB),2)
$free_memory=[Math]::Round(($ComputerMemory.FreePhysicalMemory /1MB),2)
$used_memory_perc=[math]::Round(((($Total_physical_memory - $free_memory)*100)/ $Total_physical_memory), 2)
Write-Output "<tr><td>$($VMname)</td><td>$($RebootStatus)</td><td>$($Uptime)</td><td>$($autoservice)</td><td>$($Free_perc)</td><td>$($used_memory_perc)</td><td>$($CPUavgLoad)</td></tr>"|Out-File -Append $filename       
} 
catch {
if($_.Exception.Message -like "The RPC server is unavailable.*")
{
Write-Output "<tr><td>$($VMname)</td><td>The RPC server is unavailable.</td><td></td><td></td><td></td><td></td><td></td></tr>"|Out-File -Append $filename
}
elseif($_.Exception.Message -like "Access is denied.*")
{
Write-Output "<tr><td>$($VMname)</td><td>Access is denied.</td><td></td><td></td><td></td><td></td><td></td></tr>"|Out-File -Append $filename
}
else
{
Write-Output "<tr><td>$($VMname)</td><td>$($_.Exception.Message)</td><td></td><td></td><td></td><td></td><td></td></tr>"|Out-File -Append $filename
}
}
} 

Write-Output '</table>'|Out-File -Append $filename
Write-Output '</body></html>'|Out-File -Append $filename
Write-Host "Script Ended at $([DateTime]::Now)." -BackgroundColor Black -ForegroundColor Cyan
