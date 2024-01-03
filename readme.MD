Place the contents of the repository in its own folder.  

Command syntax:

.\Densify-Export-CSV.ps1 -instance <instancename>

If you do not specify an instance name you will be prompted for one at runtime as well as a user name and password for the Densify instance.  To remember these values you can create environment variables named DensifyUser and DensifyPass respectively.  


Repository contents:

Densify-Export-CSV.ps1 - 
The actual script that pulls data from a customer instance and outputs it to a set of CSV files.

Instance Catalog.csv - 
A CSV file of all the instance types and their properties for AWS, Azure, and GCP.  New versions are periodically uploaded to https://github.com/dbc13543/InstanceCatalog/

1. Cloud and Containers.xlsb
2. Tags and Labels.xlsb

These are two Excel templates that allow you to view, slice, and dice the data output by the script.


Output files:

Each time you run the script it will create 4 output files as follows:

yyyy-mm-dd <instancename> Analyses.csv -- A list of the analyses in your Densify instance.  In most cases this is a list of accounts or subscriptions in which services were found that can be analyzed by the Densify engine.
yyyy-mm-dd <instancename> Cloud Recommendations.csv -- A list of all cloud recommendations across all CSP.
yyyy-mm-dd <instancename> Container Recommendations.csv -- A list of all container recommendations across all clusters.
yyyy-mm-dd <instancename> Tags.csv -- A raw dump of all tags and values found in all instances.  One line per tag.

Remember that you must be in a PowerShell prompt to run this script, not a command shell.  If you're new to using PowerShell and you can't run the script at all, open a PowerShell prompt as Administrator and run this command:

Set-ExecutionPolicy -ExecutionPolicy ByPass

Close and re-open your PowerShell prompt and you should be good to go!