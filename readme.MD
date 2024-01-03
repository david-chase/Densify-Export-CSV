Place the contents of the ZIP file in their own folder.  

Command syntax:

.\qryInstance -instance <instancename> [-labels]

*** Be sure to edit qryInstance.ps1 with any text editor and input your Densify user id and password here:

   [string]$user='<your username>', 
   [string]$pass='<your password>'

ZIP file contents:

qryInstance.ps1 - 
The actual script that pulls data from a customer instance and outputs it to a set of CSV files.

Template.xlsb - 
Template used to surface the results.  A copy of this file is copied and renamed each time you run the tool.  

Instance Catalog.csv - 
A CSV file of all the instance types and their properties for AWS and Azure.  GCP to follow when I have time.

Get-AppOwnerReport.ps1 - 
A work in progress.  Call this script with "-instance <instancename> -entityid <entityid>" and it will query the customer's instance using token-based auth, download the app owner report, and open it.  Sadly there's no way to pass parameters to it from Excel so it's not very useful at the moment.

Remember that you must be in a PowerShell prompt to run this script, not a command shell.  If you're new to using PowerShell and you can't run the script at all, open a PowerShell prompt as Administrator and run this command:

Set-ExecutionPolicy -ExecutionPolicy ByPass

Close and re-open your PowerShell prompt and you should be good to go!