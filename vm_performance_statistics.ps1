###
# Report that gives CPU, Memory, Disk and Network usage for a pre-defined period of days, 
# It can be performed with a schedule task and sent the report via mail to specific people.
###

# Bypass  policy 
$Bypass = Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
$Bypass

# Connect to vCenter
# Connect-VIServer <vCenter-IP-FQDN>

### Variables
# Variable in order to check how many days before we need to set the perfomance monitor 
$dates_to_check ="-10"
# Interval minutes
$Int_Min ="5"

# Find number of powered on VMs in order to create width size of the graphs later
[int]$totalvms=(Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"} | Measure-Object).Count

# Output file path in order to store the values
$Date = (Get-Date -f "ddMMyyyy")

# Path in order to store the file
$XLS_Path = 'C:\ibm_apar\AverageUsage_'+$Date+'.xlsx'
$destination = $XLS_Path
$Excel = New-Object -ComObject Excel.Application

### Create the charts 
# Create chart for CPU
$CPU = New-ExcelChart -Title 'CPU Performance' -Width 1024 -Height ($totalvms*30) -Row 0 -Column 0 -ChartType BarClustered -Header 'Usage' -SeriesHeader "CPU %" -XRange 'Usage[Name]' -YRange @('Usage[AVG CPU Usage]')

# Create chart for Memory
$Mem = New-ExcelChart -Title 'Memory Performance' -Width 1024 -Height ($totalvms*30) -Row 0 -Column 0 -ChartType BarClustered -Header 'Usage' -SeriesHeader "Memory %" -XRange 'Usage[Name]' -YRange @('Usage[AVG Memory Usage]')

# Create chart for Disk
$Disk= New-ExcelChart -Title 'Disk Performance' -Width 1024 -Height ($totalvms*30) -Row 0 -Column 0 -ChartType BarClustered -Header 'Usage' -SeriesHeader "Disk KBps" -XRange 'Usage[Name]' -YRange @('Usage[AVG Disk Usage]')

# Create chart for Network
$Net = New-ExcelChart -Title 'Network Performance' -Width 1024 -Height ($totalvms*30) -Row 0 -Column 0 -ChartType BarClustered -Header 'Usage' -SeriesHeader "Network MBps" -XRange 'Usage[Name]' -YRange @('Usage[AVG Network Usage]')

# Take values for all vms for CPU,Memory,Disk and network
Get-VM | Where {$_.PowerState -eq "PoweredOn"} | Select Name, vmhost, NumCpu, MemoryGB, `

# Take values for all vms for CPU
@{N="AVG CPU Usage"; E={[Math]::Round((($_ | Get-Stat -Stat cpu.usage.average -Start (Get-Date).AddDays($dates_to_check) -IntervalMins $Int_Min | Measure-Object Value -Average).Average),2)}}, `

# Take values for all vms for Memory
@{N="AVG Memory Usage" ; E={[Math]::Round((($_ | Get-Stat -Stat mem.usage.average -Start (Get-Date).AddDays($dates_to_check) -IntervalMins $Int_Min | Measure-Object Value -Average).Average),2)}} , `

# Take values for all vms for Network usage
@{N="AVG Network Usage" ; E={[Math]::Round((($_ | Get-Stat -Stat net.usage.average -Start (Get-Date).AddDays($dates_to_check) -IntervalMins $Int_Min | Measure-Object Value -Average).Average),2)}} , `

# Take values for all vms for Disk Usage
@{N="AVG Disk Usage" ; E={[Math]::Round((($_ | Get-Stat -Stat disk.usage.average -Start (Get-Date).AddDays($dates_to_check) -IntervalMins $Int_Min | Measure-Object Value -Average).Average),2)}} |`

# Save it in Excel file
# Create a table with all the values and vms
Export-Excel -Path $XLS_Path -WorkSheetname Usage -TableName Usage # -ExcelChartDefinition $CPU,$Mem,$Disk,$Net

# Create a graph for CPU in different sheet
Export-Excel -Path $XLS_Path -WorksheetName 'CPU_Graph' -TableName 'CPU_Graph' -ExcelChartDefinition $CPU

# Create a graph for Memory in different sheet
Export-Excel -Path $XLS_Path -WorksheetName 'Memory_Graph' -TableName 'Memory_Graph' -ExcelChartDefinition $Mem 

# Create a graph for Disk in different sheet
Export-Excel -Path $XLS_Path -WorksheetName 'Disk_Graph' -TableName 'Disk_Graph' -ExcelChartDefinition $Disk

# Create a graph for Network in different sheet
Export-Excel -Path $XLS_Path -WorksheetName 'Network_Graph' -TableName 'Network_Graph' -ExcelChartDefinition $Net

# Disconnect from vCenter
# Disconnect-viserver -confirm:$false

# Sent mail with the attachment.
# send-mailmessage -from "Perf_Mon@Customer.com" -to "joedoe@ibm.com" -subject "Perfomance Charts $(get-date -f "dd-MM-yyyy")" -body "Below you can find the rvtools report. Please see attachment `n `n `n" -Attachments $destination -smtpServer MailServerIP
