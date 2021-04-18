# Virtual Machine Performance Statistics Report

## Scope:
Script that exports perforamnce reports for all the powered on vms (CPU, Memory, Disk and Network) and creates an excel file with 5 tabs.
1. Table with all vms usage info.
2. Graph with CPU average usage (%)
3. Graph with Memory average usage (%)
4. Graph with disk average usage (KBps)
5. Graph with network average usage (MBps) 

## Requirements:
* Windows Server 2012 and above // Windows 10
* Powershell 5.1 and above
* PowerCLI either standalone or import the module in Powershell (Preferred)
* Import-Excel Module 
* MS Excel 

## Configuration
In order to set the days for the monitor change the variable $dates_to_check. By default is 10 days/
```powershell
$dates_to_check ="-10"
```

To save the excel file to some specific folder modify the following variable
```powershell
$XLS_Path = 'C:\ibm_apar\AverageUsage_'+$Date+'.xlsx'
```
For sending the result file via e-mail you will need to have the SMTP ip to be set instead of "mailServerIP", uncomment the line and set the "From", "To" and "Subject":
```
 send-mailmessage -from "Perf_Mon@Customer.com" -to "joedoe@ibm.com" -subject "Perfomance Charts $(get-date -f "dd-MM-yyyy")" -body "Below you can find the rvtools report. Please see attachment `n `n `n" -Attachments $destination -smtpServer MailServerIP
 ```


## Example
First you need to connect to the vCenter
```powershell
 PS> Connect-VIServer <vCenter-IP-FQDN>
 ```
 And then run the script
 ```powershell
 # make sure to change the directory in case you are not running the script from C:\
 PS> C:\vm_perfomance_statistics.ps1 
 ```

![Alt text](/screenshot/tab1.jpg?raw=true "Main Usage")
 
![Alt text](/screenshot/tab2.jpg?raw=true "CPU Usage")

## Frequetly Asked Questions:
* When I am executing the script it gives you an error "vCenter not found".
   > Before you execute the script you need first to be connected on a vCenter Server.
   ```powershell
   PS> Connect-VIServer <vCenter-IP-FQDN>
   ```
   
* When I run the script it gives me error on Excel commands
  > You are missing the Excel module. You need to import it prior of running the script.
  ```powershell 
  PS> Install-Module -Name ImportExcel
  ```
