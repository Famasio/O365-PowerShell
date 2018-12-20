<#
.Synopsis
   HIGH BUSINESS IMPACT - USER LICENSE/SERVICE MODIFICATION
   Function Created per User Request
   Single subscription disablement
.DESCRIPTION
   This script will view and edit all user sku's.

   Please see 
   https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
   for specific subscription names.
.INPUTS
   None
.OUTPUTS
   None
.NOTES
   v-bakwi
   Please note that this script has High Business Impact as it modifies user licenses/services. 
   Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   N/A
.ROLE
   User License/Service Management
.FUNCTIONALITY
   Edit user SKU to modify target subscription
#>
function Replace-Subscription
{
    [CmdletBinding()]
    Param (
       # Please enter Subscription SKU of main subscription you want to enable. EXAMPLE: -SubscriptionToAdd O365_BUSINESS_PREMIUM for Office365 Business Premium
            [Parameter(Mandatory=$true)]
            $SubscriptionToAdd,

       # Please enter Subscription SKU of main subscription you want to disable. EXAMPLE: -SubscriptionToRemove O365_BUSINESS_PREMIUM for Office365 Business Premium
            [Parameter(Mandatory=$true)]
            $SubscriptionToRemove
      
    )

    Begin {

        if ($Global:FunctionRun -eq $null) {
            install-module MSOnline
            import-module MSOnline
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
            $credential = Get-Credential
            Connect-MsolService -Credential $credential
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
            $Global:FunctionRun = $True
        }

        else {
        write-output "User currently signed in, aborting signup"
        }

    }

    Process {

        $users = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}
        $SelectedSubAdd = Get-MsolAccountSku | where {$_.accountskuid -like "*:$SubscriptionToAdd"}
        $SelectedSubRm = Get-MsolAccountSku | where {$_.accountskuid -like "*:$SubscriptionToRemove"}

            write-host ""
            write-host "Users affected: Everyone " -ForegroundColor "Cyan"
            write-host "Subscription being enabled:" -ForegroundColor "Cyan" $($SelectedSubAdd.accountskuid) 
            write-host "Subscription being disabled:" -ForegroundColor "Cyan" $($SelectedSubRM.accountskuid) 
            write-host ""

        $Conf = Read-host -Prompt "Please confirm the operation [Any:Confirm | N:Stop]"
        write-host ""

        if($Conf -eq 'n'){
        break
        }

        else{}

        foreach ($U in $Users){

            if ($U.licenses.accountskuid -like "*:$SubscriptionToRemove") {

                write-host "Processing $($U.displayname) - $($U.userprincipalname)"
                write-host ""

                write-host "Removed $($SelectedSubRm.accountskuid)"
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -RemoveLicenses $($SelectedSubRm.accountskuid)

                write-host "Added $($SelectedSubAdd.accountskuid)"
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -AddLicenses $($SelectedSubAdd.accountskuid)
            }

            else {
                write-host "Processing $($U.displayname) - $($U.userprincipalname)"
                write-host ""

                write-host "Added $($SelectedSubAdd.accountskuid)"
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -AddLicenses $($SelectedSubAdd.accountskuid)
                
            }
        }
    }

    End
    {
    write-host "All Users have been reviewed"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Replace-Subscription -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION"
write-host "Cmdlet get-msolaccountsku can be used to view company-specific SKU Names"
write-host ""
get-help Replace-Subscription -ShowWindow
