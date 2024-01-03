# Densify-Export-CSV

<img src="https://www.densify.com/wp-content/uploads/densify.png" width="250">

Place the contents of the repository in its own folder.  

## Command syntax:

.\Densify-Export-CSV.ps1 -instance <instancename>

If you do not specify an instance name you will be prompted for one at runtime as well as a user name and password for the Densify instance.  To remember these values you can create environment variables named DensifyUser and DensifyPass respectively.  

## Repository contents:

Densify-Export-CSV.ps1 - 
The actual script that pulls data from a customer instance and outputs it to a set of CSV files.

Instance Catalog.csv - 
A CSV file of all the instance types and their properties for AWS, Azure, and GCP.  New versions are periodically uploaded to https://github.com/dbc13543/InstanceCatalog/

1. Cloud and Containers.xlsb
2. Tags and Labels.xlsb
3. Software Identified.xlsb

These are two Excel templates that allow you to view, slice, and dice the data output by the script.

## Output files:

Each time you run the script it will create 4 output files as follows:

yyyy-mm-dd <instancename> Analyses.csv -- A list of the analyses in your Densify instance.  In most cases this is a list of accounts or subscriptions in which services were found that can be analyzed by the Densify engine. <br/>
yyyy-mm-dd <instancename> Cloud Recommendations.csv -- A list of all cloud recommendations across all CSP. <br/>
yyyy-mm-dd <instancename> Container Recommendations.csv -- A list of all container recommendations across all clusters. <br/>
yyyy-mm-dd <instancename> Tags.csv -- A raw dump of all tags and values found in all instances.  One line per tag. <br/>

## Notes

Remember that you must be in a PowerShell prompt to run this script, not a command shell.  If you're new to using PowerShell and you can't run the script at all, open a PowerShell prompt as Administrator and run this command:

Set-ExecutionPolicy -ExecutionPolicy ByPass

Close and re-open your PowerShell prompt and you should be good to go!

## Revision History
v3.3 - Added a calculation to multiple monthly cost by Avg Group Size for ASGs.  Added a column for Avg Group Size.
v3.4 - Renamed the program to be more descriptive.  Removed functionality to copy the Excel template.
v3.5 - Can now specify user and password using environment variables or you will be prompted.  Can no longer save them in a local file.
v4.0 - Now exports "yyy-mm-dd <instance> Software Identified.csv" to show software identified in Public Cloud.        - 
       Updated "1. Cloud and Containers" to gracefully handle CSVs that are missing the column "Avg Group Size".
       Added "3. Software Identified.xlsb" for exploring the Software Identified csv file.
