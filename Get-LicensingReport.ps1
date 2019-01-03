<#
.Synopsis
   NO BUSINESS IMPACT - User License Report Generation
   Function Created per User Request
   User License Report Generation
.DESCRIPTION
   Generate report about licenses based on parameters provided (See "Parameters Tab")

   It only pulls data from O365 Admin Center and does not modify any records.

   PLEASE SEE "SYNTAX" & "PARAMETERS" FOR CORRECT FUNCTION USAGE
   
   SERVICE SKU:
   https://docs.microsoft.com/en-us/office365/enterprise/powershell/view-licenses-and-services-with-office-365-powershell

   SUBSCRIPTION SKU:
   https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference

.EXAMPLE
    Get-LicenseReport -UnlicensedReport -FilePath "C:\Users\Public\Documents"
    Will create Unlicensed Users Report at C:\Users\Public\Documents\UnlicensedReport.csv
    
    Get-LicenseReport -UnlicensedReport -OnlyInternal -FilePath "C:\Users\Public\Documents"
    Will create Unlicensed Users (Internal Only) Report at C:\Users\Public\Documents\UnlicensedReport.csv

    Get-LicenseReport -SubscriptionReport -FilePath "C:\Users\Public\Documents"
    Will create User Subscription Report at C:\Users\Public\Documents\SubscriptionReport.csv

    Get-LicenseReport -DetailedReport -FilePath "C:\Users\Public\Documents"
    Will create User Service Name & Status Report at C:\Users\Public\Documents\DetailedReport.csv

    Get-LicenseReport -UsageReport -FilePath "C:\Users\Public\Documents"
    Will create Subscription Usage Report (Available/Assigned) at C:\Users\Public\Documents\UsageReport.csv

    Get-LicenseReport -All -FilePath "C:\Users\Public\Documents"
    Will create all 4 reports in C:\Users\Public\Documents folder.

    Get-LicenseReport -Filter -FilterList '$SubscriptionID -like "FLOW_FREE" -and $ServiceName[$i] -like "Teams1" -and $ServiceStatus[$i] -like "Success"' -FileName "Flow + Teams Active.csv" -FilePath "C:\Users\Public\Documents"
    Will list all users, who have Flow Free assigned, and in addition have Teams Service in active state


.INPUTS
   -Filepath "Your filepath to FOLDER"
.OUTPUTS
   "UnlicensedReport.csv" "SubscriptionReport.csv" "DetailedReport.csv" "UsageReport.csv"
.NOTES
   v-bakwi
   Please note that this script has No Business Impact as it does not modify any settings. 
   Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   MSOnline Module / EXO Session Import (Both Self-Contained)
.ROLE
   User License/Service Reporting
#>
function Get-LicenseReport
{
    [CmdletBinding()]
    Param (
        # Will create a report consisting of users that are unlicensed
        [Parameter(Mandatory=$true, ParameterSetName = 'Unlicensed')]
        [Switch]$UnlicensedReport,

        # Excludes EXT Users & Guests from Unlicensed Report
        [Parameter(ParameterSetName = 'Unlicensed')]
        [Switch]$OnlyInternal,
       
        # Will create a report consisting of users and their subscription plans
        [Parameter(Mandatory=$true, ParameterSetName = 'Subscription')]
        [Switch]$SubscriptionReport,

        # Will create a report consisting of users and their detailed services (SKU's)
        [Parameter(Mandatory=$true, ParameterSetName = 'SKUDetails')]
        [Switch]$DetailedReport,

        # Will create a report consisting of all subscriptions on tenant, licenses available and assigned
        [Parameter(Mandatory=$true, ParameterSetName = 'UsageReport')]
        [Switch]$UsageReport,

        # Will generate all reports (ie. Unlicensed, Subscription, Detailed & Usage at provided location)
        [Parameter(Mandatory=$true, ParameterSetName = 'All')]
        [Switch]$All,
        
        # Will enable use of custom filters in order to generate personalized CSV file about users
        [Parameter(Mandatory=$true, ParameterSetName = 'Filtered')]
        [Switch]$Filter,

        # List of used filters in format: -FilterList '$SubscriptionID -like "FLOW_FREE" -and $ServiceName[$i] -like "Teams1" -and $ServiceStatus[$i] -like "Success"'. $SubscriptionID is String ID from 2nd link, $ServiceName[$i] and $ServiceStatus[$i] HAVE to be used in exact same form with $i. ServiceName is listed in 1st link, where ServiceStatus can be "Success, PendingInput, Disabled" (May be other)
        [Parameter(Mandatory=$true, ParameterSetName = 'Filtered')]
        $FilterList,

        # Makes you able to add custom filename to personalized reports, like: -FileName "Flow + Teams Active.csv"
        [Parameter(Mandatory=$true, ParameterSetName = 'Filtered')]
        $Filename,

        # Filepath to folder you want your report(s) in, like -FilePath "C:\Users\Public\Documents"
        [Parameter(Mandatory=$true, ParameterSetName = 'All')]
        [Parameter(Mandatory=$true, ParameterSetName = 'Unlicensed')]
        [Parameter(Mandatory=$true, ParameterSetName = 'Subscription')]
        [Parameter(Mandatory=$true, ParameterSetName = 'SKUDetails')]
        [Parameter(Mandatory=$true, ParameterSetName = 'UsageReport')]
        [Parameter(Mandatory=$true, ParameterSetName = 'Filtered')]
        $FilePath
        )

    Begin {

        if ($Global:FunctionRun -eq $null) {

            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
            Install-Module MSOnline -ErrorAction Stop
            Import-Module MsOnline -ErrorAction Stop
            $credential = Get-Credential
            Connect-MsolService -Credential $credential -ErrorAction Stop
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" –AllowRedirection -ErrorAction Stop  
            Import-PSSession $ExchangeSession -AllowClobber -ErrorAction Stop
            $Global:FunctionRun = $True
            cls
        }

        else {
        cls
        write-host "User currently signed in, aborting signup"
        write-host ""
        }

        if ($All){

            $UnlicensedReport = $true
            $SubscriptionReport = $true
            $DetailedReport = $true
            $UsageReport = $true
        }

    }

    Process {
        
        if ($UnlicensedReport){

            write-host ""
            write-host "Generating report about unlicensed users" -foregroundcolor "Cyan"
            write-host ""

            $headerstringmain = "UserPrincipalName,DisplayName"
            Out-File -FilePath $FilePath\UnlicensedReport.csv -InputObject $headerstringmain -Encoding UTF8 -append 

            if ($OnlyInternal){

                $Users  = Get-Msoluser -All -Unlicensedusersonly | where {$_.userprincipalname -notlike "*#EXT#*"}
            }
            else {

                $Users  = Get-Msoluser -All -Unlicensedusersonly
            }

            foreach ($U in $Users) {
                write-host "Processing $($U.displayname)"
                $headerstring = ($U.displayname + "," + $U.UserPrincipalName) 
                Out-File -FilePath $FilePath\UnlicensedReport.csv -InputObject $headerstring -Encoding UTF8 -append 

            }
        }

        if ($SubscriptionReport){

            write-host ""
            write-host "Generating report about user subscriptions" -foregroundcolor "Cyan"
            write-host ""
            
            $dummystring = ""
            $headerstringmain = "DisplayName,UserPrincipalName,SubscriptionName"
            Out-File -FilePath $FilePath\SubscriptionReport.csv -InputObject $headerstringmain -Encoding UTF8 -append 

            $Users  = Get-Msoluser -All | Where-Object {$_.isLicensed -eq $true}
            
            foreach ($U in $Users) {
                
                Out-File -FilePath $FilePath\SubscriptionReport.csv -InputObject $dummystring -Encoding UTF8 -append

                write-host "Processing $($U.displayname)"

                $AccountSKUID = $AccountSKUID = $U.licenses.accountskuid.split(":") | Select -unique
                $AccountSKUID = $AccountSKUID[1..($AccountSKUID.length-1)]
                $headersubstring = ($U.displayname + "," + $U.UserPrincipalName + ",") 

                Out-File -FilePath $FilePath\SubscriptionReport.csv -InputObject $headersubstring -Encoding UTF8 -append

                foreach ($SKU in $AccountSKUID) {

                    $headersubstring = ("," + "," + $SKU) 
                    Out-File -FilePath $FilePath\SubscriptionReport.csv -InputObject $headersubstring -Encoding UTF8 -append
                }
                

            }
        }

        if ($DetailedReport){

            write-host ""
            write-host "Generating detailed report about user services/SKU's" -foregroundcolor "Cyan"
            write-host ""

            $dummystring = ""
            $headerstringmain = "DisplayName,UserPrincipalName,ServiceName,ProvisioningStatus"
            Out-File -FilePath $FilePath\DetailedReport.csv -InputObject $headerstringmain -Encoding UTF8 -append 

            $Users  = Get-Msoluser -All | Where-Object {$_.isLicensed -eq $true}
            
            foreach ($U in $Users) {
                
                Out-File -FilePath $FilePath\DetailedReport.csv -InputObject $dummystring -Encoding UTF8 -append

                write-host "Processing $($U.displayname)"

                $AccountSKUName = $u.Licenses.servicestatus.serviceplan.servicename
                $AccountSKUStatus = $u.Licenses.servicestatus.provisioningstatus
                $headersubstring = ($U.displayname + "," + $U.UserPrincipalName + ",") 

                Out-File -FilePath $FilePath\DetailedReport.csv -InputObject $headersubstring -Encoding UTF8 -append

                for($i = 0; $i -lt $AccountSKUName.count; $i++){

                    $headersubstring = ("," + "," +  $AccountSKUName[$i] + ","+ $AccountSKUStatus[$i]) 
                    Out-File -FilePath $FilePath\DetailedReport.csv -InputObject $headersubstring -Encoding UTF8 -append
                }
                

            }

        }

        if ($UsageReport){

            write-host ""
            write-host "Generating report about subscription usage" -foregroundcolor "Cyan"
            write-host ""

            $SKUs = Get-MsolAccountSku

            $dummystring = ("Subscription SKU" + "," + "Available" + "," + "Assigned")
            Out-File -FilePath $FilePath\UsageReport.csv -InputObject $dummystring -Encoding UTF8 -append

            ForEach ($S in $SKUs) {

                write-host "Processing: " $($S.AccountSkuId.split(":")[1])

                $headersubstring = ($($S.AccountSkuId.split(":")[1]) + "," + $($S.ActiveUnits) + "," + $($S.ConsumedUnits)) 
                Out-File -FilePath $FilePath\UsageReport.csv -InputObject $headersubstring -Encoding UTF8 -append
            }
        }

        if ($Filter){

            write-host ""
            write-host "Generating custom report with filter:" $FilterList -foregroundcolor "Cyan"
            write-host ""

            $Users  = Get-Msoluser -All | Where-Object {$_.isLicensed -eq $true}

            $sb = [scriptblock]::create($FilterList)

            $headermainstring = ("DisplayName" + "," + "UserPrincipalName" + "," + "UsageLocation" + ",") 
            Out-File -FilePath $FilePath\$FileName -InputObject $headermainstring -Encoding UTF8 -append

            foreach ($U in $Users) {
                
                write-Host ""
                write-host "Processing $($U.displayname)"

                $ServiceName = $u.Licenses.servicestatus.serviceplan.servicename

                $ServiceStatus = $u.Licenses.servicestatus.provisioningstatus

                $SubscriptionID = $SubscriptionID = $u.licenses.accountskuid.split(":") | Select -unique
                $SubscriptionID = $SubscriptionID[1..($SubscriptionID.length-1)]

                for($i = 0; $i -lt $ServiceName.count; $i++){
                    
                    if ($U | where $sb) {

                        $headersubstring = ($U.displayname + "," + $U.UserPrincipalName + "," + $U.UsageLocation + ",") 
                        Out-File -FilePath $FilePath\$FileName -InputObject $headersubstring -Encoding UTF8 -append
                        write-Host "User meets filter - added to CSV file" -ForegroundColor Cyan
                        continue

                    }
                    else {

                    }
                    
                }

            Start-Sleep -Milliseconds 250
            }

        }

    }

    End
    {
    write-host "DONE" -foregroundcolor "Cyan"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Get-LicenseReport -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION" -foregroundcolor "Cyan"
write-host "PLEASE SEE https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference FOR SKU ID REFERENCE" -foregroundcolor "Cyan"
write-host ""
get-help Get-LicenseReport -ShowWindow
