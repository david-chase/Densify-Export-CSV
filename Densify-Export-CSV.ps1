#-----------------------------------------------------------------------------------------------
#  Densify-Export-CSV v3.4
#  PowerShell script to query a customer instance and export relevant information as a CSV 
# 
#  v3.3 - Added a calculation to multiple monthly cost by Avg Group Size for ASGs.  Added a column for Avg Group Size.
#  v3.4 - Renamed the program to be more descriptive.  Removed functionality to copy the Excel template.
#------------------------------------------------------------------------------------------------

param (
   [string]$instance="",
   [string]$user='dchase@densify.com', 
   [string]$pass='oM2(hV(mcfBUj7wn',
   [Int32]$sleep = 0,
   [switch]$beep = $false
)

Write-Host -----------------------------------------------------------
Write-Host " "qryInstance - Summarize data for a customer instance
Write-Host -----------------------------------------------------------

$sPassword = ConvertTo-SecureString -String $pass -AsPlainText -Force
$oCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $sPassword

# Halt if no instance was specified on command-line or the config file
if( -not  $instance ) { Write-Host "No instance specified..." -ForegroundColor Red; exit 1 }

# Instances can be specified as fully qualified or in short form.  If user specified it in short form, add the .densify.com part
# If they specified it in long form, just parse out the short form and store it in $sInstanceTitle.  We'll use this later
if( -not $instance.Contains( ".densify.com" ) ) {
    $sInstanceTitle = $instance
    $instance += ".densify.com" 
} # END if
else {
    $sInstanceTitle = $instance.Substring( 0, $instance.IndexOf( "." ) )
} # END else

# Setup some strings
$sBaseUri = "https://" + $instance + "/CIRBA/api/v2/"
$hHeaders = @{ "Accept"="application/json" }

$aEndPoints = @( "aws", "azure", "gcp" )
# $aEndPoints = @( "azure" )
$aResourceTags = @()
Write-Output "Using instance $sInstanceTitle"

#
# Public Cloud
#

# Loop through each supported cloud name
foreach( $sEndPoint in $aEndPoints ) {

    Write-Output "Processing $sEndPoint" 
    # Write-Output "Hitting endpoint $sUri"
    try { 
        # Generate a token
        $sUri = $sBaseUri + "authorize"
        $hBody = @{ 'userName'=$user; 'pwd'=$pass }
        $hHeaders = @{ "Accept"="application/json"; "Content-Type"="application/json" }
        $oAuthToken = Invoke-RestMethod -Method Post -Uri $sUri -Headers $hHeaders -Body ( $hBody | ConvertTo-Json ) -ErrorAction Stop
        # Hit the web API endpoint
        $hHeaders = @{ "Authorization"= [System.String]::Concat( "Bearer ", $oAuthToken.apiToken  );"Accept"="application/json" }
        $sUri = $sBaseUri + "analysis/cloud/" + $sEndPoint
        $aAnalyses = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop 
    } catch {
        # An error was thrown
        switch( $_.Exception.Response.StatusCode ) {
            "Unauthorized" { 
                Write-Output "ERROR 401: Unauthorized"
                Write-Output "Edit the script to check your credentials"
                Exit
            } # switch case
            default {
                Write-Output "ERROR unhandled exception occurred"
                Exit
            } # switch case

        } # switch( $_.Exception.Response.StatusCode.value__ )
          
    } # catch
    $aAllAnalyses += $aAnalyses
    $iCurrentAnalysis = 0

    # Loop through each cloud analysis
    foreach( $oAnalysis in $aAnalyses ) {
        $iCurrentAnalysis++
        
        # If a sleep parameter was passed on the command line wait that number of seconds.  No need to sleep on the very first iteration.
        if( ( $sleep -gt 0 ) -and ( $iCurrentAnalysis -gt 1 ) ) {
            Write-Output "Sleeping for $sleep seconds"
            Start-Sleep -Seconds $sleep
        } # if( $sleep -gt 0 )

        # Add a custom column for AccountFriendlyName
        if( $oAnalysis.accountName ) { $sFriendlyName = $oAnalysis.accountName + " (" + $oAnalysis.accountId + ")" } else { $sFriendlyName = "(" + $oAnalysis.accountId + ")" }
        $oAnalysis | add-member -name "Account Friendly Name" -type NoteProperty -value $sFriendlyName
        
        # Convert UNIX epoch times to Windows DateTime
        $oAnalysis.analysisCompletedOn = $oAnalysis.analysisCompletedOn / 1000 / 86400 + 25569

        # Write-Output "Hitting endpoint $sUri"
        $sOutString = [System.String]::Concat( "Processing analysis [", $iCurrentAnalysis, "/", $aAnalyses.Count, "] in ", $sEndPoint )
        Write-Output $sOutString
        try { 
            $sUri = $sBaseUri + "authorize"
            $hBody = @{ 'userName'=$user; 'pwd'=$pass }
            $hHeaders = @{ "Accept"="application/json"; "Content-Type"="application/json" }
            $oAuthToken = Invoke-RestMethod -Method Post -Uri $sUri -Headers $hHeaders -Body ( $hBody | ConvertTo-Json ) -ErrorAction Stop
            # Hit the web API endpoint
            $hHeaders = @{ "Authorization"= [System.String]::Concat( "Bearer ", $oAuthToken.apiToken  );"Accept"="application/json" }
            $sUri = $sBaseUri + "analysis/cloud/" + $sEndPoint + "/" + $oAnalysis.analysisID + "/results?includeAttributes=true"
            $aRecos = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop 
        } catch {
        # An error was thrown
            switch( $_.Exception.Response.StatusCode ) {
                "Unauthorized" { 
                    Write-Output "ERROR 401: Unauthorized"
                    Write-Output "Edit the script to check your credentials"
                    Exit
                } # switch case
                default {
                    Write-Output "ERROR unhandled exception occurred"
                    Exit
                } # switch case
            } # switch( $_.Exception.Response.StatusCode )
        } # catch
       
        # Loop through all Recommendations
        foreach( $oReco in $aRecos ) { 

            # Add a list of Attributes
            $oReco | add-member -name "Account Friendly Name" -type NoteProperty -value ""
            $oReco | add-member -name "Memory Data Source" -type NoteProperty -value ""
            $oReco | add-member -name "Mem Util (%)" -type NoteProperty -value ""
            $oReco | add-member -name "Mem Util (%) Limit" -type NoteProperty -value ""
            $oReco | add-member -name "CPU Util (%)" -type NoteProperty -value ""
            $oReco | add-member -name "CPU Util (%) Limit" -type NoteProperty -value ""
            $oReco | add-member -name "Network IO (Bytes)" -type NoteProperty -value ""
            $oReco | add-member -name "Network IO (Bytes) Limit" -type NoteProperty -value ""
            $oReco | add-member -name "Sizing Notes" -type NoteProperty -value ""
            $oReco | add-member -name "Business Unit" -type NoteProperty -value ""
            $oReco | add-member -name "Application" -type NoteProperty -value ""
            $oReco | add-member -name "Owner" -type NoteProperty -value ""
            $oReco | add-member -name "Operational Environment" -type NoteProperty -value ""
            $oReco | add-member -name "Inventory Code" -type NoteProperty -value ""
            $oReco | add-member -name "Department" -type NoteProperty -value ""
            $oReco | add-member -name "Project" -type NoteProperty -value ""
            $oReco | add-member -name "Cost Center" -type NoteProperty -value ""
            $oReco | add-member -name "Product Code" -type NoteProperty -value ""
            $oReco | add-member -name "Existing CPU Allocation" -type NoteProperty -value ""
            $oReco | add-member -name "Existing CPU Benchmark" -type NoteProperty -value ""
            $oReco | add-member -name "ENA Support" -type NoteProperty -value ""
            $oReco | add-member -name "Existing Memory Allocation" -type NoteProperty -value ""
            $oReco | add-member -name "Cloud Provider" -type NoteProperty -value ""
            $oReco | add-member -name "Config Last Changed On" -type NoteProperty -value ""
            $oReco | add-member -name "mswcName" -type NoteProperty -value ""
            $oReco | add-member -name "Kubernetes Cluster" -type NoteProperty -value ""

            # If this is an ASG, multiply the current costs by the average instance count
            if( $oReco."serviceType" -eq "ASG" ) {
                $oReco.currentCost = $oReco.currentCost * $oReco.avgInstanceCountCurrent
            } # END if( $oReco."serviceType" -eq "ASG" )

            # Populate the field "Cloud Provider". 
            switch( $oReco."serviceType" ) {
                "ASG" { $oReco."Cloud Provider" = "AWS" }
                "Compute Engine" { $oReco."Cloud Provider" = "GCP" }
                "EC2" { $oReco."Cloud Provider" = "AWS" }
                "RDS" { $oReco."Cloud Provider" = "AWS" }
                "SPOT" { $oReco."Cloud Provider" = "AWS" }
                "Virtual Machine" { $oReco."Cloud Provider" = "Azure" }
                "Scale Set" { $oReco."Cloud Provider" = "Azure" } # This is wild guess what the service type will be named
                default { $oReco."Cloud Provider" = "Unrecognized Service Type" }
            } # switch( $oReco.SizingNotes )

            # Add the account friendly name
            $aMatchingAccounts = $aAnalyses | where { $_.accountId -eq $oReco.accountIdRef }
            $oReco."Account Friendly Name" = $aMatchingAccounts[ 0 ]."Account Friendly Name"

            # Convert UNIX epoch times to Windows DateTime
            if( $oReco.recommFirstSeen ) { $oReco.recommFirstSeen = ( $oReco.recommFirstSeen / 1000 ) / 86400 + 25569 }
            if( $oReco.recommLastSeen ) { $oReco.recommLastSeen = ( $oReco.recommLastSeen / 1000 ) / 86400 + 25569 }

            # Include Predicted Uptime costs just like the UI does
            if( $oReco."currentCost" ) { $oReco."currentCost" = $oReco."currentCost" * $oReco."predictedUptime" / 100 }
            if( $oReco."recommendedCost" ) { $oReco."recommendedCost" = $oReco."recommendedCost" * $oReco."predictedUptime" / 100 }

            # Fully qualify the app owner report link
            # if( $oReco."rptHref" ) { $oReco."rptHref" = $sBaseUri + $oReco."rptHref".SubString( 1, $oReco."rptHref".Length - 1 ) }
            
            foreach( $oAttribute in $oReco.Attributes ) { 
                switch ( $oAttribute.name ) {
                    "Mem Util (%)" {
                        $oReco."Mem Util (%)" = $oAttribute.value
                     } # switch case
                    "Mem Util (%) Limit" {
                        $oReco."Mem Util (%) Limit" = $oAttribute.value
                     } # switch case
                    "CPU Util (%)" {
                        $oReco."CPU Util (%)" = $oAttribute.value
                     } # switch case
                    "CPU Util (%) Limit" {
                        $oReco."CPU Util (%) Limit" = $oAttribute.value
                     } # switch case
                    "Network IO (Bytes)" {
                        $oReco."Network IO (Bytes)" = $oAttribute.value
                     } # switch case
                    "Network IO (Bytes) Limit" {
                        $oReco."Network IO (Bytes) Limit" = $oAttribute.value
                    } # switch case
                    "Sizing Notes" {
                        $oReco."Sizing Notes" = $oAttribute.value
                     } # switch case
                    "Business Unit" {
                        $oReco."Business Unit" = $oAttribute.value
                     } # switch case
                    "Application" {
                        $oReco."Application" = $oAttribute.value
                     } # switch case
                    "Owner" {
                        $oReco."Owner" = $oAttribute.value
                     } # switch case
                    "Operational Environment" {
                        $oReco."Operational Environment" = $oAttribute.value
                     } # switch case
                    "Inventory Code" {
                        $oReco."Inventory Code" = $oAttribute.value
                     } # switch case
                    "Department" {
                        $oReco."Department" = $oAttribute.value
                     } # switch case
                    "Project" {
                        $oReco."Project" = $oAttribute.value
                     } # switch case
                    "Cost Center" {
                        $oReco."Cost Center" = $oAttribute.value
                     } # switch case
                    "Product Code" {
                        $oReco."Product Code" = $oAttribute.value
                     } # switch case
                    "Existing CPU Allocation" {
                        $oReco."Existing CPU Allocation" = $oAttribute.value
                     } # switch case
                    "Existing CPU Benchmark" {
                        $oReco."Existing CPU Benchmark" = $oAttribute.value
                     } # switch case
                    "ENA Support" {
                        $oReco."ENA Support" = $oAttribute.value
                     } # switch case
                    "Existing Memory Allocation" {
                        $oReco."Existing Memory Allocation" = $oAttribute.value
                     } # switch case
                    "Config Last Changed On" {
                        $oReco."Config Last Changed On" = $oAttribute.value
                     } # switch case
                    "mswcName" {
                        $oReco."mswcName" += ( $oAttribute.value.SubString( $oAttribute.value.IndexOf( ":" ) + 1, $oAttribute.value.Length - $oAttribute.value.IndexOf( ":" ) - 1 )  ) + ", "
                     } # switch case
                    "Kubernetes Cluster" {
                        $oReco."Kubernetes Cluster" = $oAttribute.value
                     } # switch case
                    "Resource Tags" {
                        $oResourceTag = [PSCustomObject]@{
                            entityId = $oReco."entityId"
                            "System Name" = $oReco."name"
                            "Cloud Provider" = $oReco."Cloud Provider"
                            "Tag Name" = $oAttribute.value.Split( " : ", 2 )[ 0 ]
                            "Value" = $oAttribute.value.Split( " : ", 2 )[ 1 ]
                        }
                        $aResourceTags += $oResourceTag
                     } # switch case
                } #switch ( $oAttribute.name )

                # Populate the field "Memory Data Source". 
                switch( $oReco."Sizing Notes" ) {
                    "Sized via actual workload data" { $oReco."Memory Data Source" = "Actual" }
                    "No Memory Utilization Data" { $oReco."Memory Data Source" = "None/Backfill" }
                    "Sized via fallback workload" { $oReco."Memory Data Source" = "None/Backfill" }
                } # switch( $oReco.SizingNotes )

            } # foreach( $oAttribute in $oReco )

            $oReco."mswcName" = $oReco."mswcName".TrimEnd( ", " )

        } # foreach( $oReco in $aRecos )

        $aAllRecos += $aRecos

    } # foreach( $oAnalysis in $aAnalyses )

} # ForEach( $sEndPoint in $aEndPoints )

#
# Kubernetes
#

Write-Output "Processing containers"
try {
    $sUri = $sBaseUri + "authorize"
    $hBody = @{ 'userName'=$user; 'pwd'=$pass }
    $hHeaders = @{ "Accept"="application/json"; "Content-Type"="application/json" }
    $oAuthToken = Invoke-RestMethod -Method Post -Uri $sUri -Headers $hHeaders -Body ( $hBody | ConvertTo-Json ) -ErrorAction Stop
    # Hit the web API endpoint
    $hHeaders = @{ "Authorization"= [System.String]::Concat( "Bearer ", $oAuthToken.apiToken  );"Accept"="application/json" }
    $sUri = $sBaseUri + "analysis/containers/kubernetes/"
    $aAnalyses = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop
    # $aAnalyses = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop -Authentication Basic -Credential $oCredential
} catch {
# An error was thrown
    switch( $_.Exception.Response.StatusCode ) {
        "Unauthorized" { 
            Write-Output "ERROR 401: Unauthorized"
            Write-Output "Edit the script to check your credentials"
            Exit
        } # switch case
        default {
            Write-Output "ERROR unhandled exception occurred"
            Exit
        } # switch case
    } # switch( $_.Exception.Response.StatusCode )
} # catch

$iCurrentAnalysis = 0

# Loop through each cluster analysis
foreach( $oAnalysis in $aAnalyses ) {
    $iCurrentAnalysis++
    $sOutString = [System.String]::Concat( "Processing cluster [", $iCurrentAnalysis, "/", $aAnalyses.Count, "]" )
    Write-Output $sOutString

    # If a sleep parameter was passed on the command line wait that number of seconds.  No need to sleep on the very first iteration.
    if( ( $sleep -gt 0 ) -and ( $iCurrentAnalysis -gt 1 ) ) {
        Write-Output "Sleeping for $sleep seconds"
        Start-Sleep -Seconds $sleep
    } # if( $sleep -gt 0 )

    $sUri = $sBaseUri + "authorize"
    $hBody = @{ 'userName'=$user; 'pwd'=$pass }
    $hHeaders = @{ "Accept"="application/json"; "Content-Type"="application/json" }
    $oAuthToken = Invoke-RestMethod -Method Post -Uri $sUri -Headers $hHeaders -Body ( $hBody | ConvertTo-Json ) -ErrorAction Stop
    # Hit the web API endpoint
    $hHeaders = @{ "Authorization"= [System.String]::Concat( "Bearer ", $oAuthToken.apiToken  );"Accept"="application/json" }
    $sUri = $sBaseUri + "analysis/containers/kubernetes/" + $oAnalysis.analysisID + "/results"
    $aRecos = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop
    # $aRecos = Invoke-RestMethod -Method Get -Uri $sUri -Headers $hHeaders -ErrorAction Stop -Authentication Basic -Credential $oCredential

    # Loop through all Recommendations
    foreach( $oReco in $aRecos ) { 
            # Convert UNIX epoch times to Windows DateTime
            if( $oReco.recommFirstSeen ) { $oReco.recommFirstSeen = ( $oReco.recommFirstSeen / 1000 ) / 86400 + 25569 }
            if( $oReco.recommLastSeen ) { $oReco.recommLastSeen = ( $oReco.recommLastSeen / 1000 ) / 86400 + 25569 }
    } # foreach( $oReco in $aRecos )

    $aAllContainerRecos += $aRecos

} # foreach( $oAnalysis in $aAnalyses )

#
# Post processing and write to disk
#

# Remove, reorder, and rename cloud recommendations columns
$aAllRecos = $aAllRecos | Select-Object @{ N="entityId"; E={ $_.entityId } }, 
    @{ N="Account ID"; E={ $_.accountIdRef } }, 
    @{ N="Account Friendly Name"; E={ $_."Account Friendly Name" } }, 
    @{ N="Resource ID"; E={ $_.resourceId } }, 
    @{ N="System Name"; E={ $_.name } },
    @{ N="Service Type"; E={ $_.serviceType } },
    @{ N="Cloud Provider"; E={ $_."Cloud Provider" } },
    @{ N="Optimization Type"; E={ $_.recommendationType } },
    @{ N="Current Instance Type"; E={ $_.currentType } },  
    @{ N="Recom. Instance Type"; E={ $_.recommendedType } },
    @{ N="Effort"; E={ $_.effortEstimate } },
    @{ N="Predicted Uptime"; E={ $_.predictedUptime } },
    @{ N="Current Monthly Cost"; E={ $_.currentCost } },
    @{ N="Recom. Monthly Cost"; E={ $_.recommendedCost } },
    @{ N="Monthly Savings Estimate"; E={ $_.savingsEstimate } },
    @{ N="Current Hourly Rate"; E={ $_.currentHourlyRate } },
    @{ N="Recom. Hourly Rate"; E={ $_.recommendedHourlyRate } },
    @{ N="Total Hours Running"; E={ $_.totalHoursRunning } },
    @{ N="Implementation Method"; E={ $_.implementationMethod } },
    @{ N="Recommendation First Seen"; E={ $_.recommFirstSeen } },
    @{ N="Recommendation Last Seen"; E={ $_.recommLastSeen } },
    @{ N="Times Recommendation Seen"; E={ $_.recommSeenCount } },
    @{ N="RI Coverage %"; E={ $_.currentRiCoverage } },
    @{ N="Defer Recommendation"; E={ $_.deferRecommendation } },
    @{ N="Densify Policy"; E={ $_.densifyPolicy } },
    @{ N="Mem Util (%)"; E={ $_."Mem Util (%)" } },
    @{ N="Mem Util (%) Limit"; E={ $_."Mem Util (%) Limit" } },
    @{ N="CPU Util (%)"; E={ $_."CPU Util (%)" } },
    @{ N="CPU Util (%) Limit"; E={ $_."Mem Util (%) Limit" } },
    @{ N="Network IO (Bytes)"; E={ $_."Network IO (Bytes)" } },
    @{ N="Network IO (Bytes) Limit"; E={ $_."Network IO (Bytes) Limit" } },
    @{ N="Sizing Notes"; E={ $_."Sizing Notes" } },
    @{ N="Memory Data Source"; E={ $_."Memory Data Source" } },
    @{ N="Current CPU Count"; E={ $_."Existing CPU Allocation" } },
    @{ N="Current CPU Benchmark"; E={ $_."Existing CPU Benchmark" } },
    @{ N="ENA Support"; E={ $_."ENA Support" } },
    @{ N="Current Memory Amount"; E={ $_."Existing Memory Allocation" } },
    @{ N="Instance Type Last Changed On"; E={ $_."Config Last Changed On" } },
    @{ N="Software Identified"; E={ $_."mswcName" } },
    @{ N="Application"; E={ $_."Application" } },
    @{ N="Business Unit"; E={ $_."Business Unit" } },
    @{ N="Owner"; E={ $_."Owner" } },
    @{ N="Operational Environment"; E={ $_."Operational Environment" } },
    @{ N="Inventory Code"; E={ $_."Inventory Code" } },
    @{ N="Department"; E={ $_."Department" } },
    @{ N="Project"; E={ $_."Project" } },
    @{ N="Cost Center"; E={ $_."Cost Center" } },
    @{ N="Product Code"; E={ $_."Product Code" } },
    # @{ N="AWS Region"; E={ $_.region } },
    # @{ N="App Owner Report"; E={ $_.rptHref } },
    # @{ N="Approval Type"; E={ $_.approvalType } },
    @{ N="Power State"; E={ $_.powerState } },
    @{ N="Kubernetes Cluster"; E={ $_."Kubernetes Cluster" } },
    @{ N="Avg Group Size"; E={ $_.avgInstanceCountCurrent } }

    # @{ N="recommendedHostEntityId"; E={ $_.recommendedHostEntityId } },
    # @{ N="Audit Info"; E={ $_.AuditInfo } },

# Remove, reorder, and rename container recommendations columns
# "currentCount"
$aAllContainerRecos = $aAllContainerRecos | Select-Object @{ N="entityId"; E={ $_.entityId } },
    @{ N="Namespace"; E={ $_.namespace } },
    @{ N="Cluster"; E={ $_.cluster } },
    @{ N="Host Name"; E={ $_.hostName } },
    @{ N="Container"; E={ $_.container } },
    @{ N="Display Name"; E={ $_.displayName } },
    @{ N="Pod Service"; E={ $_.podService } },
    # @{ N="Audit Info"; E={ $_.auditInfo } },
    @{ N="Controller Type"; E={ $_."controllerType" } },
    @{ N="# of Containers"; E={ $_."currentCount" } },
    @{ N="Optimization Type"; E={ $_."recommendationType" } },
    @{ N="Predicted Uptime"; E={ $_."predictedUptime" } },
    @{ N="Current CPU Request"; E={ $_."currentCpuRequest" } },
    @{ N="Recom. CPU Request"; E={ $_."recommendedCpuRequest" } },
    @{ N="Current CPU Limit"; E={ $_."currentCpuLimit" } },
    @{ N="Recom. CPU Limit"; E={ $_."recommendedCpuLimit" } },
    @{ N="Current Memory Request"; E={ $_."currentMemRequest" } },
    @{ N="Recom. Memory Request"; E={ $_."recommendedMemRequest" } },
    @{ N="Current Memory Limit"; E={ $_."currentMemLimit" } },
    @{ N="Recom. Memory Limit"; E={ $_."recommendedMemLimit" } },
    # ALERT!!!  Test if estimated savings needs to have predicted uptime factored in
    @{ N="Estimated Monthly Savings"; E={ $_."estimatedSavings" } },
    @{ N="Recommendation First Seen"; E={ $_.recommFirstSeen } },
    @{ N="Recommendation Last Seen"; E={ $_.recommLastSeen } },
    @{ N="Times Recommendation Seen"; E={ $_."recommSeenCount" } }
    

# Decide on all my output file names
$sAnalysesFile = "$PSScriptRoot" + "\" + ( Get-Date -Format "yyyy-MM-dd" ) + " " + $sInstanceTitle + " Analyses.csv"
$sCloudFile = "$PSScriptRoot" + "\" + ( Get-Date -Format "yyyy-MM-dd" ) + " " + $sInstanceTitle + " Cloud Recommendations.csv"
$sTagsFile = "$PSScriptRoot" + "\" + ( Get-Date -Format "yyyy-MM-dd" ) + " " + $sInstanceTitle + " Tags.csv"
$sContainersFile = "$PSScriptRoot" + "\" + ( Get-Date -Format "yyyy-MM-dd" ) + " " + $sInstanceTitle + " Container Recommendations.csv"
# $sExcelTemplateSource = $PSScriptRoot + "\Template.xlsb" 
# $sExcelTemplateTarget = "$PSScriptRoot" + "\" + ( Get-Date -Format "yyyy-MM-dd" ) + " " + $sInstanceTitle + " Instance Export.xlsb"

# Dump results to a file
$sOutString = "Writing " + $aAllAnalyses.Count + " analyses"
Write-Output $sOutString
$aAllAnalyses | ConvertTo-Csv | Out-File -FilePath $sAnalysesFile

$sOutString = "Writing " + $aAllRecos.Count + " cloud recommendations"
Write-Output $sOutString
$aAllRecos | ConvertTo-Csv | Out-File -FilePath $sCloudFile

$sOutString = "Writing " + $aAllContainerRecos.Count + " container recommendations"
Write-Output $sOutString
# Output only headers if the containers recos is blank
if( $aAllContainerRecos.Count -eq 0 ) { 
    $sOutstring = "entityId,Namespace,Cluster,Host Name,Container,Display Name,Pod Service,Controller Type,# of Containers,Optimization Type,Predicted Uptime,Current CPU Request,Recom. CPU Request,Current CPU Limit,Recom. CPU Limit,Current Memory Request,Recom. Memory Request,Current Memory Limit,Recom. Memory Limit,Estimated Monthly Savings,Recommendation First Seen,Recommendation Last Seen,Times Recommendation Seen"
    $sOutstring | Out-File -FilePath $sContainersFile 
    $sOutstring = ",,,,,,,,,,,,,,,,,,,,,,"
    $sOutstring | Out-File -FilePath $sContainersFile -Append
}
else {
    $aAllContainerRecos | ConvertTo-Csv | Out-File -FilePath $sContainersFile
}

$sOutString = "Writing list of " + $aResourceTags.Count + " tags and labels"
Write-Output $sOutString
$aResourceTags | ConvertTo-Csv | Out-File -FilePath $sTagsFile
# Write-Output "Writing Excel file"
# Copy-Item $sExcelTemplateSource $sExcelTemplateTarget

if( $beep ) { [console]::beep(200,500) }