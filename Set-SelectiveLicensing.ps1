<#
.Synopsis
   HIGH BUSINESS IMPACT - User License Bulk Modification
   Function Created per User Request
   User License Modification
.DESCRIPTION
   Modify licenses for all users ina  group
   Articles containing Plan & SKU ID's and Abbrev. for User's reference.
   https://docs.microsoft.com/en-us/office365/enterprise/powershell/view-licenses-and-services-with-office-365-powershell
   https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
.INPUTS
   N/A
.OUTPUTS
   N/A
.NOTES
   v-bakwi
   Please note that this script has HIGH BUSINESS IMPACT and has to be handled with caution. 
   Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   MSOnline Module / EXO Session Import / SharePoint Mgmt Module (All Self-Contained)
.ROLE
   User Bulk License Edit
#>
function Set-SelectiveLicensing
{
    [CmdletBinding()]
    Param (

        # Use when setting licenses for a specific group
         [Parameter(Mandatory=$true, ParameterSetName = 'Group')]
        [Switch]$Group,

        # Use when setting licenses for whole tenant
         [Parameter(Mandatory=$true, ParameterSetName = 'Bulk')]
        [Switch]$Bulk,

        # Use to import custom CSV Database with users to be modified (CSV Database requires "UserPrincipalName Header and list of UPN's below")
         [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
        [Switch]$Import,

         # Use to export all licensed users on tenant into a CSV File
         [Parameter(Mandatory=$true, ParameterSetName = 'Export')]
        [Switch]$Export,

        # Please enter Subscription SKU of main subscription you want to assign. EXAMPLE: -Subscription O365_BUSINESS_PREMIUM for Office365 Business Premium (WITHOUT "CompanyName:")
         [Parameter(Mandatory=$true, ParameterSetName = 'Group')]
         [Parameter(Mandatory=$true, ParameterSetName = 'Bulk')]
         [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
        $Subscription,

        # Please enter Service SKU's of set subscription that you want disabled. EXAMPLE: -DisabledServices "FLOW_O365_P1,MCOSTANDARD,POWERAPPS_O365_P1,SHAREPOINTWAC"
         [Parameter(Mandatory=$true, ParameterSetName = 'Group')]
         [Parameter(Mandatory=$true, ParameterSetName = 'Bulk')]
         [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
        $DisabledServices,
       
        # Will modify all users from this group. EXAMPLE: -GroupName TestGroup or -GroupName "Group with spaces" (Exact name from O365 Portal)
        [Parameter(Mandatory=$true, ParameterSetName = 'Group')]
        $GroupName,


        # Filepath to CSV file you have your users in
        [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
        [Parameter(Mandatory=$true, ParameterSetName = 'Export')]
        $FilePath
        )

    Begin {

        if ($Global:FunctionRun -eq $null) {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
            $credential = Get-Credential

            if(!(import-module msonline)){
                Install-Module Msonline
                Import-Module Msonline
            }
            Connect-MsolService -Credential $credential
            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
            Import-PSSession $exchangeSession
            $Global:FunctionRun = $True
        }

        else {
        }
    }

    Process {

        if($Group) {

            $SelectedSub = Get-MsolAccountSku | where {$_.accountskuid -like "*:$Subscription"} #$($SelectedSub.accountskuid)
            $SelectedGroup =  Get-MsolGroupMember -GroupObjectId “$GroupName”
            $SelectedDisableArr = $($DisabledServices -split ",")

            write-host ""
            write-host "Selected Group: "$GroupName -ForegroundColor "Cyan"
            write-host "Selected Subscription Name: "$($SelectedSub.accountskuid) -ForegroundColor "Cyan"
            write-host "Disabled services:"$SelectedDisableArr -ForegroundColor "Cyan"

            $Conf = Read-host -Prompt "Please confirm the operation [Any:Confirm | N:Stop]"
            write-host ""

            if($Conf -eq 'n'){
                break
            }

            else{}
        
            forEach ($GroupUser in $($SelectedGroup.emailaddress)){

                write-host "Processing $Groupuser from $GroupName"

                $NewPlan = New-MsolLicenseOptions -AccountSkuId $SelectedSub -DisabledPlans $SelectedDisableArr
                Set-MsolUserLicense -UserPrincipalName $Groupuser -LicenseOptions $NewPlan

            }
        }

        if($Bulk){
        
            $users = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}
            $SelectedSubAdd = Get-MsolAccountSku | where {$_.accountskuid -like "*:$Subscription"}
            $SelectedDisableArr = $($DisabledServices -split ",")

            write-host ""
            write-host "Selected Users: Everyone " -ForegroundColor "Cyan"
            write-host "Subscription being enabled:" -ForegroundColor "Cyan" $($SelectedSubAdd.accountskuid) 
            write-host "Disabled services:"$SelectedDisableArr -ForegroundColor "Cyan" 
            write-host ""

            $Conf = Read-host -Prompt "Please confirm the operation [Any:Confirm | N:Stop]"
            write-host ""

            if($Conf -eq 'n'){
                break
            }

            else{}

            foreach ($U in $Users){

                write-host "Processing $($U.displayname) - $($U.userprincipalname)"
                write-host ""

                $NewPlan = New-MsolLicenseOptions -AccountSkuId $SelectedSub -DisabledPlans $SelectedDisableArr
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -LicenseOptions $NewPlan
            }

        }

        if($Import){

            $users = Import-Csv -Path $FilePath
            $SelectedSubAdd = Get-MsolAccountSku | where {$_.accountskuid -like "*:$Subscription"}
            $SelectedDisableArr = $($DisabledServices -split ",")

            write-host ""
            write-host "Selected: CSV File " -ForegroundColor "Cyan"
            write-host "Subscription being enabled:" -ForegroundColor "Cyan" $($SelectedSubAdd.accountskuid) 
            write-host "Disabled services:"$SelectedDisableArr -ForegroundColor "Cyan" 
            write-host ""

            $Conf = Read-host -Prompt "Please confirm the operation [Any:Confirm | N:Stop]"
            write-host ""

            if($Conf -eq 'n'){
                break
            }

            else{}

            foreach ($U in $Users){

                write-host "Processing $($U.userprincipalname)"
                write-host ""

                $NewPlan = New-MsolLicenseOptions -AccountSkuId $SelectedSub -DisabledPlans $SelectedDisableArr
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -LicenseOptions $NewPlan
            }


        }

        if($Export){

            $users = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true} | Select userprincipalname
            
            $users | Export-Csv -Path "$FilePath\UserExport.csv"

        }

    }

    End
    {
    write-host "DONE" -foregroundcolor "Cyan"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Set-SelectiveLicensing -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION" -foregroundcolor "Cyan"
write-host "PLEASE SEE https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference FOR SKU ID REFERENCE" -foregroundcolor "Cyan"
write-host ""
get-help Set-SelectiveLicensing -ShowWindow
