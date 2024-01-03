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



v4.0<br/>
---<br/>
* Now exports "yyy-mm-dd <instance> Software Identified.csv" to show software identified in Public Cloud.  <br/>      - 
* Updated "1. Cloud and Containers" to gracefully handle CSVs that are missing the column "Avg Group Size".<br/>
* Added "3. Software Identified.xlsb" for exploring the Software Identified csv file.<br/>

v3.5<br/>
----<br/>
* Can now specify user and password using environment variables or you will be prompted.  Can no longer save them in a local file.<br/>

v3.4<br/>
----<br/>
* Renamed the program to be more descriptive.  <br/>
* Removed functionality to copy the Excel template.<br/>

v3.3<br/>
-----<br/>
* Added a calculation to multiple monthly cost by Avg Group Size for ASGs.  <br/>
* Added a column for Avg Group Size.<br/>

v3.2<br/>
----<br/>
* Reverted back to token-based auth
* Abandoned container labels.  The amount of data it pulls -- even in a small environment -- is prohibitive.  Doesn't even work with the subscription client so no point trying to make it work here
* Added an instance type catalog.  Only includes AWS and Azure at the moment.  This surfaces a lot of details about the current and recommended instance type that previously weren't there
* Added the following fields to the Excel template that are all based on the instance catalog: Current Instance Category (Memory optimized, Compute optimized, General Purpose etc), Recom. Instance Category, Current Instance Platform (Intel/AMD/ARM), Recom. Instance Platform, Current Instance Arch. Year (Year the *oldest* chipset in this instance type was released), Recom. Instance Arch. Year, Current Instance Info (Web link to the CSP web site with instance tech details), Recom. Instance Info
* Added the field Instance Type Last Changed On which helps identify when a customer is implementing recommendations.
* Added the field Software Identified which corresponds to mswcName in DCE.  This field lists any software that was identified when running the Software Identification internal job.  Note that this field will be blank if a POP report has never been run, and will only be as current as the last time PM ran the Software Identification job.

v3.1<br/>
----<br/>
* Container labels are now working, but they take a VERY long time to collect.<br/>
* You must add the switch -labels at the end of your command-line if you want to parse container labels.  By default it will skip container labels.<br/>

v3.0<br/>
----<br/>
* Rewrote the script and Excel file from scratch to be simpler and faster.<br/>
* Now exports Public Cloud attributes so you can slice and dice based on the cloud tags you've mapped for your customer.
* No longer uses a config file.  Just edit the .ps1 file to change your username and password.<br/>

v2.1<br/>
----<br/>
* No need to update data source anymore.  It just assumes the data files are named using the same convention as the .xlsm file.<br/>
* Reordered the columns in the Cloud tab for Brett.  Hi Brett!  :) <br/>
* In the Cloud Tab, the columns named "Current Monthly Cost" and "Recommended Monthly Cost" now INCLUDE predicted uptime. 
 This fixes the long-standing discrepancy between what's reported by the API versus what's reported by the UI.<br/>

v2.0<br/>
----<br/>
* Generates a new authentication token before each query.  This prevents timeouts for large instances<br/>
* Added a "filtering" tab to the spreadsheet so you can automatically apply some "low-hanging fruit" filters<br/>
* Updated the "Cloud" and "Containers" tables to be essentially customer-ready.  This means columns were given friendly names, unnecessary columns were hidden, dates were converted from Unix epoch time to Windows Date/Time.  Used a join query so account names are shown in a friendly format "(Number) Long name" instead of just a number.<br/>
* Added support for container labels using Attributes.  This slows down processing but makes it significantly more powerful.<br/>
* No longer uses a Subscription to collect Cloud Tags, uses Attributes instead.<br/>
* Can now ingest Attributes while querying the API.  Ingesting all attributes slows processing to a crawl so only supports certain ones.  Right now it ingests CPU Util %, Memory Util %, Resource (Cloud) Tags, Container Labels, Sizing Notes, and Container Info.<br/>
* Added a tab to spreadsheet that surfaces Attributes by EntityID.  Note that EntityID is the primary key used to link Attributes to data in all other tables.<br/>
* Added a config tab so it's easy to change your data source.  Just edit it on the config tab instead of manually modifying a bunch of queries.<br/>
* Added a Tags & Labels tab that features two Pivot Tables for summaries of cloud tags and container labels.<br/>

v0.91<br/>
------<br/>
* Removed a limit of no more than 3,000 supported Resource Tags.  Now supports up to 30,000.<br/>
* The program now outputs a third .CSV file named "Accounts.csv".  This is a list of all the AWS accounts, Azure subscriptions, or GCP projects that have been ingested into this subscription including their friendly names.  I added a new tab to the "qryInstance Analysis.xlsx" file that includes this data.<br/>
* If you specify CopyExcelTemplate=true in qryInstance.config it will look for a file named "qryInstance Analysis.template.xlsx in the script folder and rename it to match the naming format of your .csv output files.  So if your output files are named "2022-12-23 acmecorp *.csv" it will create a file called "2022-12-23 acmecorp Instance Output.xlsx"  This is only to save you some work renaming .xlsx files, you still need to update the data sources manually.  This is turned off in the default config file.<br/>
