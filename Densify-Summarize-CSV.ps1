#-----------------------------------------------------------------------------------------------
#  Densify-Summarize-CSV v0.9
#  PowerShell script to examine a bunch of extracted CSVs and summarize them
# 
#  v0.9 - Test version
#------------------------------------------------------------------------------------------------

param (
   [string]$instance="",
   [string]$output = ""
)

# Return the right "X" characters of a string
function RightStr {
    param( 
        [Parameter( Mandatory )]
        [string]$sString,
        [Parameter( Mandatory )]
        [int]$iIndex
    )

        return $sString.Substring( $sString.Length - 1, $iIndex )
} # END function RightStr

function AddTrailBackslash {
    param( 
        [Parameter( Mandatory )]
        [string]$sString
    )

    if( ( RightStr $sString 1 ) -ne "\" ) { return $sString + "\" } else { return $sString }
} # END function AddTrailBackslash

Write-Host 
Write-Host ::: Densify-Summarize-CSV ::: -ForegroundColor Cyan
Write-Host 

# Suss out the user instance, and output folder
if( -not $user ) { $user = $env:DensifyUser }
if( -not $output -and $env:DensifyOutput ) { $output = AddTrailBackslash -sString $env:DensifyOutput }
if( -not $instance ) { $instance = Read-Host -Prompt "Enter instance name" }
if( -not $output ) { $output = AddTrailBackslash -sString $PSScriptRoot }

# Set the filespec
$sFileSpec = [System.String]::Concat( $output, "????-??-?? ", $instance, " Cloud Recommendations.csv"  )
$aCSVFiles = @( Get-ChildItem $sFileSpec -File )

# Initialize the array of properties
$aCSVSummaries = @()

# Define the name of the summary file
$sOutFile = [System.String]::Concat( $output, $instance, " Summary.CSV" )

# Run a counter for status
$iCurrentCSV = 0

foreach( $sCSVFile in $aCSVFiles ) {

    # Show some status
    $iCurrentCSV++
    Write-Progress -Activity "Parsing CSV Files" -Status "Progress:" -PercentComplete ( $iCurrentCSV / $aCSVFiles.Count * 100 )

    $sCSVPrefix = [System.String]::Concat( $output, $sCSVFile.Name.Split( " " )[ 0 ], " ", $instance, " " )
    $sCSVPostfix = ".csv"
    $sExportDate = ( [Datetime]::ParseExact( $sCSVFile.Name.Split( " " )[ 0 ], 'yyyy-MM-dd', $null ) ).ToString( "dd-MMM-yyyy")

    # Process the Analyses file
    $aCSVFile = Import-CSV -Path ( $sCSVPrefix + "Analyses" + $sCSVPostfix )
    $aAnalysisCount = $aCSVFile.Count
    $iAWS_Analysis_Count = ( $aCSVFile | Where-Object analysisResults -CLike "*/aws/*" ).Count
    $iAzure_Analysis_Count = ( $aCSVFile | Where-Object analysisResults -CLike "*/azure/*" ).Count
    $iGCP_Analysis_Count = ( $aCSVFile | Where-Object analysisResults -CLike "/*gcp/*" ).Count

    #EC2
    # PUBLIC CLOUD
    #

    # Read in the Cloud Recommendations file
    $aCSVLines = Import-CSV -Path ( $sCSVFile )

    # Process nothing if the CSV file is empty
    if( $aCSVLines.Count -gt 0 ) {
        # EC2
        $aEC2 = $aCSVLines | Where-Object { ( $_."Cloud Provider" -eq "AWS" ) -and ( $_."Service Type" -eq "EC2" ) -and ( $_."Optimization Type" -ne "Not Analyzed" ) }
        $iEC2_Count = $aEC2.Count
        $iEC2_JustRight_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Just Right" ) } ).Count
        $iEC2_Downsize_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Downsize" ) } ).Count
        $iEC2_DownsizeOptimal_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Downsize - Optimal Family" ) } ).Count
        $iEC2_Upsize_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Upsize" ) } ).Count
        $iEC2_UpsizeOptimal_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Upsize - Optimal Family" ) } ).Count
        $iEC2_Modernize_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Modernize" ) } ).Count
        $iEC2_Terminate_Count = ( $aEC2 | Where-Object { ( $_."Optimization Type" -eq "Terminate" ) } ).Count
        # If there are no objects of this type then Measure-Object will fail
        if( $iEC2_Count -gt 0 ) {
            $iEC2_Current_Cost = $aEC2 | Measure-Object -Property "Current Monthly Cost" -Sum | Select-Object -ExpandProperty "Sum"
            $iEC2_Savings = $aEC2 | Measure-Object -Property "Monthly Savings Estimate" -Sum | Select-Object -ExpandProperty "Sum"
            $iEC2_Savings_Pct = $iEC2_Savings / $iEC2_Current_Cost * 100
        } else {
            $iEC2_Current_Cost = 0
            $iEC2_Savings = 0
            $iEC2_Savings_Pct = 0
        } # END if( $iEC2_Count -gt 0 )

        # RDS
        $aRDS = $aCSVLines | Where-Object { ( $_."Cloud Provider" -eq "AWS" ) -and ( $_."Service Type" -eq "RDS" ) -and ( $_."Optimization Type" -ne "Not Analyzed" ) }
        $iRDS_Count = $aRDS.Count
        $iRDS_JustRight_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Just Right" ) } ).Count
        $iRDS_Downsize_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Downsize" ) } ).Count
        $iRDS_DownsizeOptimal_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Downsize - Optimal Family" ) } ).Count
        $iRDS_Upsize_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Upsize" ) } ).Count
        $iRDS_UpsizeOptimal_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Upsize - Optimal Family" ) } ).Count
        $iRDS_Modernize_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Modernize" ) } ).Count
        $iRDS_Terminate_Count = ( $aRDS | Where-Object { ( $_."Optimization Type" -eq "Terminate" ) } ).Count
        # If there are no objects of this type then Measure-Object will fail
        if( $iRDS_Count -gt 0 ) {
            $iRDS_Current_Cost = $aRDS | Measure-Object -Property "Current Monthly Cost" -Sum | Select-Object -ExpandProperty "Sum"
            $iRDS_Savings = $aRDS | Measure-Object -Property "Monthly Savings Estimate" -Sum | Select-Object -ExpandProperty "Sum"
            $iRDS_Savings_Pct = $iRDS_Savings / $iRDS_Current_Cost * 100
        } else {
            $iRDS_Current_Cost = 0
            $iRDS_Savings = 0
            $iRDS_Savings_Pct = 0
        } # END if( $iRDS_Count -gt 0 )            

        # ASG
        $aASG = $aCSVLines | Where-Object { ( $_."Cloud Provider" -eq "AWS" ) -and ( $_."Service Type" -eq "ASG" ) -and ( $_."Optimization Type" -ne "Not Analyzed" ) }
        $iASG_Count = $aASG.Count
        $iASG_JustRight_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Just Right" ) } ).Count
        $iASG_Downsize_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Downsize" ) } ).Count
        $iASG_DownsizeOptimal_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Downsize - Optimal Family" ) } ).Count
        $iASG_Upsize_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Upsize" ) } ).Count
        $iASG_UpsizeOptimal_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Upsize - Optimal Family" ) } ).Count
        $iASG_Modernize_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Modernize" ) } ).Count
        $iASG_Terminate_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Terminate" ) } ).Count
        $iASG_Downscale_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Downscale" ) } ).Count
        $iASG_Upscale_Count = ( $aASG | Where-Object { ( $_."Optimization Type" -eq "Upscale" ) } ).Count
        # If there are no objects of this type then Measure-Object will fail
        if( $iASG_Count -gt 0 ) {
            $iASG_Current_Cost = $aASG | Measure-Object -Property "Current Monthly Cost" -Sum | Select-Object -ExpandProperty "Sum"
            $iASG_Savings = $aASG | Measure-Object -Property "Monthly Savings Estimate" -Sum | Select-Object -ExpandProperty "Sum"
            $iASG_Savings_Pct = $iASG_Savings / $iASG_Current_Cost * 100
            # Clean up the bug in the data.  For wild numbers set ASG savings percent to zero
            if( $iASG_Savings_Pct -ge 10000 ) { $iASG_Savings_Pct = 0 } # "∞"
        } else {
            $iASG_Current_Cost = 0
            $iASG_Savings = 0
            $iASG_Savings_Pct = 0
        } # END if( $iASG_Count -gt 0 )     

        # Virtual Machines
        $aAzVM = $aCSVLines | Where-Object { ( $_."Cloud Provider" -eq "Azure" ) -and ( $_."Service Type" -eq "Virtual Machine" ) -and ( $_."Optimization Type" -ne "Not Analyzed" ) }
        $iAzVM_Count = $aAzVM.Count
        $iAzVM_JustRight_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Just Right" ) } ).Count
        $iAzVM_Downsize_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Downsize" ) } ).Count
        $iAzVM_DownsizeOptimal_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Downsize - Optimal Family" ) } ).Count
        $iAzVM_Upsize_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Upsize" ) } ).Count
        $iAzVM_UpsizeOptimal_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Upsize - Optimal Family" ) } ).Count
        $iAzVM_Modernize_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Modernize" ) } ).Count
        $iAzVM_Terminate_Count = ( $aAzVM | Where-Object { ( $_."Optimization Type" -eq "Terminate" ) } ).Count
        # If there are no objects of this type then Measure-Object will fail
        if( $iAzVM_Count -gt 0 ) {
            $iAzVM_Current_Cost = $aAzVM | Measure-Object -Property "Current Monthly Cost" -Sum | Select-Object -ExpandProperty "Sum"
            $iAzVM_Savings = $aAzVM | Measure-Object -Property "Monthly Savings Estimate" -Sum | Select-Object -ExpandProperty "Sum"
            $iAzVM_Savings_Pct = $iAzVM_Savings / $iAzVM_Current_Cost * 100
        } else {
            $iAzVM_Current_Cost = 0
            $iAzVM_Savings = 0
            $iAzVM_Savings_Pct = 0
        } # END if( $iAzVM_Count -gt 0 )        

        # Compute Engine
        $aCE = $aCSVLines | Where-Object { ( $_."Cloud Provider" -eq "GCP" ) -and ( $_."Service Type" -eq "Compute Engine" ) -and ( $_."Optimization Type" -ne "Not Analyzed" ) }
        $iCE_Count = $aCE.Count
        $iCE_JustRight_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Just Right" ) } ).Count
        $iCE_Downsize_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Downsize" ) } ).Count
        $iCE_DownsizeOptimal_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Downsize - Optimal Family" ) } ).Count
        $iCE_Upsize_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Upsize" ) } ).Count
        $iCE_UpsizeOptimal_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Upsize - Optimal Family" ) } ).Count
        $iCE_Modernize_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Modernize" ) } ).Count
        $iCE_Terminate_Count = ( $aCE | Where-Object { ( $_."Optimization Type" -eq "Terminate" ) } ).Count
        # If there are no objects of this type then Measure-Object will fail
        if( $iCE_Count -gt 0 ) {
            $iCE_Current_Cost = $aCE | Measure-Object -Property "Current Monthly Cost" -Sum | Select-Object -ExpandProperty "Sum"
            $iCE_Savings = $aCE | Measure-Object -Property "Monthly Savings Estimate" -Sum | Select-Object -ExpandProperty "Sum"
            $iCE_Savings_Pct = $iCE_Savings / $iCE_Current_Cost * 100
        } else {
            $iCE_Current_Cost = 0
            $iCE_Savings = 0
            $iCE_Savings_Pct = 0
        } # END if( $iCE_Count -gt 0 )

        $iWeekofYear = [int]( "{0:d1}" -f ( $( Get-Culture ).Calendar.GetWeekOfYear( ( [Datetime]$sExportDate ),[System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday ) ) )

        # Output the summary for this date
        $oCSVSummary = New-Object -TypeName PSObject -Property @{
            Date = $sExportDate
            "Week of Year" = [System.String]::Concat( $sCSVFile.Name.Split( "-" )[ 0 ], "-", $iWeekofYear.ToString( "00" ) )
            "Total Analysis Count" = $aAnalysisCount
            "AWS Analysis Count" = $iAWS_Analysis_Count
            "Azure Analysis Count" = $iAzure_Analysis_Count
            "GCP Analysis Count" = $iGCP_Analysis_Count
            "Kubernetes Analysis Count" = 0

            "AWS Services Count" = $iEC2_Count + $iRDS_Count + $iASG_RDS_Count
            "Azure Services Count" = $iAzVM_Count
            "GCP Services Count" = $iCE_Count
            "Kubernetes Services Count" = 0

            "EC2 Count" = $iEC2_Count
            "EC2 Just Right Count" = $iEC2_JustRight_Count
            "EC2 Downsize Count" = $iEC2_Downsize_Count
            "EC2 Downsize Optimal Family Count" = $iEC2_DownsizeOptimal_Count
            "EC2 Upsize Count" = $iEC2_Upsize_Count
            "EC2 Upsize Optimal Family Count" = $iEC2_UpsizeOptimal_Count
            "EC2 Modernize Count" = $iEC2_Modernize_Count
            "EC2 Terminate Count" = $iEC2_Terminate_Count
            "EC2 Savings Opport. Count" = $iEC2_Downsize_Count + $iEC2_DownsizeOptimal_Count + $iEC2_Modernize_Count + $iEC2_Terminate_Count
            "EC2 At Risk Count" = $iEC2_Upsize_Count + $iEC2_UpsizeOptimal_Count
            "EC2 Current Cost" = [math]::Round( $iEC2_Current_Cost, 2 )
            "EC2 Savings ($)" = [math]::Round( $iEC2_Savings, 2 )
            "EC2 Savings (%)" = [math]::Round( $iEC2_Savings_Pct, 1 )
            
            "RDS Count" = $iRDS_Count
            "RDS Just Right Count" = $iRDS_JustRight_Count
            "RDS Downsize Count" = $iRDS_Downsize_Count
            "RDS Downsize Optimal Family Count" = $iRDS_DownsizeOptimal_Count
            "RDS Upsize Count" = $iRDS_Upsize_Count
            "RDS Upsize Optimal Family Count" = $iRDS_UpsizeOptimal_Count
            "RDS Modernize Count" = $iRDS_Modernize_Count
            "RDS Terminate Count" = $iRDS_Terminate_Count
            "RDS Savings Opport. Count" = $iRDS_Downsize_Count + $iRDS_DownsizeOptimal_Count + $iRDS_Modernize_Count + $iRDS_Terminate_Count
            "RDS At Risk Count" = $iRDS_Upsize_Count + $iRDS_UpsizeOptimal_Count
            "RDS Current Cost" = [math]::Round( $iRDS_Current_Cost, 2 )
            "RDS Savings ($)" = [math]::Round( $iRDS_Savings, 2 )
            "RDS Savings (%)" = [math]::Round( $iRDS_Savings_Pct, 1 )
            
            "ASG Count" = $iASG_Count
            "ASG Just Right Count" = $iASG_JustRight_Count
            "ASG Downsize Count" = $iASG_Downsize_Count
            "ASG Downsize Optimal Family Count" = $iASG_DownsizeOptimal_Count
            "ASG Upsize Count" = $iASG_Upsize_Count
            "ASG Upsize Optimal Family Count" = $iASG_UpsizeOptimal_Count
            "ASG Modernize Count" = $iASG_Modernize_Count
            "ASG Terminate Count" = $iASG_Terminate_Count
            "ASG Downscale Count" = $iASG_Downscale_Count
            "ASG Upscale Count" = $iASG_Upscale_Count
            "ASG Savings Opport. Count" = $iASG_Downsize_Count + $iASG_DownsizeOptimal_Count + $iASG_Modernize_Count + $iASG_Terminate_Count + $iASG_Downscale_Count
            "ASG At Risk Count" = $iASG_Upsize_Count + $iASG_UpsizeOptimal_Count + $iASG_Upscale_Count
            "ASG Current Cost" = [math]::Round( $iASG_Current_Cost, 2 )
            "ASG Savings ($)" = [math]::Round( $iASG_Savings, 2 )
            "ASG Savings (%)" = [math]::Round( $iASG_Savings_Pct, 1 )

            "Azure VM Count" = $iAzVM_Count
            "Azure VM Just Right Count" = $iAzVM_JustRight_Count
            "Azure VM Downsize Count" = $iAzVM_Downsize_Count
            "Azure VM Downsize Optimal Family Count" = $iAzVM_DownsizeOptimal_Count
            "Azure VM Upsize Count" = $iAzVM_Upsize_Count
            "Azure VM Upsize Optimal Family Count" = $iAzVM_UpsizeOptimal_Count
            "Azure VM Modernize Count" = $iAzVM_Modernize_Count
            "Azure VM Terminate Count" = $iAzVM_Terminate_Count
            "Azure VM Savings Opport. Count" = $iAzVM_Downsize_Count + $iAzVM_DownsizeOptimal_Count + $iAzVM_Modernize_Count + $iAzVM_Terminate_Count
            "Azure VM At Risk Count" = $iAzVM_Upsize_Count + $iAzVM_UpsizeOptimal_Count
            "Azure VM Current Cost" = [math]::Round( $iAzVM_Current_Cost, 2 )
            "Azure VM Savings ($)" = [math]::Round( $iAzVM_Savings, 2 )
            "Azure VM Savings (%)" = [math]::Round( $iAzVM_Savings_Pct, 1 )

            "GCP Compute Engine Count" = $iCE_Count
            "GCP Compute Engine Just Right Count" = $iCE_JustRight_Count
            "GCP Compute Engine Downsize Count" = $iCE_Downsize_Count
            "GCP Compute Engine Downsize Optimal Family Count" = $iCE_DownsizeOptimal_Count
            "GCP Compute Engine Upsize Count" = $iCE_Upsize_Count
            "GCP Compute Engine Upsize Optimal Family Count" = $iCE_UpsizeOptimal_Count
            "GCP Compute Engine Modernize Count" = $iCE_Modernize_Count
            "GCP Compute Engine Terminate Count" = $iCE_Terminate_Count
            "GCP Compute Engine Savings Opport. Count" = $iCE_Downsize_Count + $iCE_DownsizeOptimal_Count + $iCE_Modernize_Count + $iCE_Terminate_Count
            "GCP Compute Engine At Risk Count" = $iCE_Upsize_Count + $iCE_UpsizeOptimal_Count
            "GCP Compute Engine Current Cost" = [math]::Round( $iCE_Current_Cost, 2 )
            "GCP Compute Engine Savings ($)" = [math]::Round( $iCE_Savings, 2 )
            "GCP Compute Engine Savings (%)" = [math]::Round( $iCE_Savings_Pct, 1 )

            "Container Manifest Count" = 0
            "Container Manifest Just Right Count" = 0
            "Container Manifest Downsize Count" = 0
            "Container Manifest Upsize Count" = 0
            "Container Manifest Size from Unspecified" = 0
            "Container Manifest Terminate Count" = 0
            "Container Manifest Savings Opport. Count" = 0
            "Container Manifest At Risk Count" = 0
            "Container Manifest Current Cost" = 0
            "Container Manifest Savings ($)" = 0
            "Container Manifest Savings (%)" = 0
        } # New-Object

        $aCSVSummaries += $oCSVSummary

    } # END if( $aCSVLines.Count -gt 0 )

} # END foreach( $sCSVFile in $aCSVFiles )

$sOutString = "Writing " + $aCSVSummaries.Count + " records to " + $sOutFile
Write-Output $sOutString
$aCSVSummaries | Select-Object -Property Date, "Week of Year", "Total Analysis Count", "AWS Analysis Count", "Azure Analysis Count", "GCP Analysis Count", "Kubernetes Analysis Count",  `
                "AWS Services Count", "Azure Services Count", "GCP Services Count", "Kubernetes Services Count",  `
                "EC2 Count", "RDS Count", "ASG Count", "Azure VM Count", "GCP Compute Engine Count", "Container Manifest Count", `
                "EC2 Just Right Count", "EC2 Downsize Count", "EC2 Downsize Optimal Family Count", "EC2 Upsize Count", "EC2 Upsize Optimal Family Count", `
                "EC2 Modernize Count", "EC2 Terminate Count", "EC2 Savings Opport. Count", "EC2 At Risk Count", "EC2 Current Cost", "EC2 Savings ($)", "EC2 Savings (%)", `
                "RDS Just Right Count", "RDS Downsize Count", "RDS Downsize Optimal Family Count", "RDS Upsize Count", "RDS Upsize Optimal Family Count", `
                "RDS Modernize Count", "RDS Terminate Count", "RDS Savings Opport. Count", "RDS At Risk Count", "RDS Current Cost", "RDS Savings ($)", "RDS Savings (%)", `
                "ASG Just Right Count", "ASG Downsize Count", "ASG Downsize Optimal Family Count", "ASG Upsize Count", "ASG Upsize Optimal Family Count", `
                "ASG Modernize Count", "ASG Terminate Count", "ASG Downscale Count", "ASG Upscale Count", `
                "ASG Savings Opport. Count", "ASG At Risk Count", "ASG Current Cost", "ASG Savings ($)", "ASG Savings (%)", `
                "Azure VM Just Right Count", "Azure VM Downsize Count", "Azure VM Downsize Optimal Family Count", "Azure VM Upsize Count", "Azure VM Upsize Optimal Family Count", `
                "Azure VM Modernize Count", "Azure VM Terminate Count", "Azure VM Savings Opport. Count", "Azure VM At Risk Count", "Azure VM Current Cost", "Azure VM Savings ($)", `
                "Azure VM Savings (%)", `
                "GCP Compute Engine Just Right Count", "GCP Compute Engine Downsize Count", "GCP Compute Engine Downsize Optimal Family Count", "GCP Compute Engine Upsize Count", "GCP Compute Engine Upsize Optimal Family Count", `
                "GCP Compute Engine Modernize Count", "GCP Compute Engine Terminate Count", "GCP Compute Engine Savings Opport. Count", "GCP Compute Engine At Risk Count", "GCP Compute Engine Current Cost", "GCP Compute Engine Savings ($)", `
                "GCP Compute Engine Savings (%)", `
                "Container Manifest Just Right Count", "Container Manifest Downsize Count", "Container Manifest Upsize Count", "Container Manifest Size from Unspecified", `
                "Container Manifest Terminate Count", "Container Manifest Savings Opport. Count", "Container Manifest At Risk Count", "Container Manifest Current Cost", "Container Manifest Savings ($)", `
                "Container Manifest Savings (%)" `
                | Export-CSV -Path $sOutFile -UseQuotes AsNeeded
             

                