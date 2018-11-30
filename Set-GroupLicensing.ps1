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
function Set-GroupLicensing
{
    [CmdletBinding()]
    Param (
        # Please enter Subscription SKU of main subscription you want to assign. EXAMPLE: -Subscription O365_BUSINESS_PREMIUM for Office365 Business Premium
         [Parameter(Mandatory=$true)]
        $Subscription,

        # Please enter Service SKU's that you want disabled. EXAMPLE: -DisabledServices "FLOW_O365_P1,MCOSTANDARD,POWERAPPS_O365_P1,SHAREPOINTWAC"
         [Parameter(Mandatory=$true)]
        $DisabledServices,
       
        # Will modify all users from this group. EXAMPLE: -GroupName TestGroup or -GroupName "Group with spaces" (Exact name from O365 Portal)
        [Parameter(Mandatory=$true)]
        $GroupName
        )

    Begin {

        if ($Global:FunctionRun -eq $null) {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
            $credential = Get-Credential
            Import-Module Msolservice
            Connect-MsolService -Credential $credential
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
            Import-PSSession $exchangeSession
            $Global:FunctionRun = $True
        }

        else {

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
        }
    }

    Process {

        forEach ($GroupUser in $($SelectedGroup.emailaddress)){

            

            write-host "Processing $Groupuser from $GroupName"

            $NewPlan = New-MsolLicenseOptions -AccountSkuId $SelectedSub -DisabledPlans $SelectedDisableArr
            Set-MsolUserLicense -UserPrincipalName $Groupuser -LicenseOptions $NewPlan

        }
    }

    End
    {
    write-host "DONE" -foregroundcolor "Cyan"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Set-GroupLicensing -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION" -foregroundcolor "Cyan"
write-host "PLEASE SEE https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference FOR SKU ID REFERENCE" -foregroundcolor "Cyan"
write-host ""
get-help Set-GroupLicensing -ShowWindow