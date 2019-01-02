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
            [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
            [Parameter(Mandatory=$true, ParameterSetName = 'Standard')]
            $SubscriptionToAdd,

       # Please enter Subscription SKU of main subscription you want to disable. EXAMPLE: -SubscriptionToRemove O365_BUSINESS_PREMIUM for Office365 Business Premium
            [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
            [Parameter(Mandatory=$true, ParameterSetName = 'Standard')]
            $SubscriptionToRemove,

       # Please specify path to your CSV file when using "ImportCSV" Param. Format: "C:\Something\More Something\File.csv"
            [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
            $FilePath,

       # Use to import custom CSV Database with users to be modified (CSV Database requires "UserPrincipalName Header and list of UPN's below")
            [Parameter(Mandatory=$true, ParameterSetName = 'Import')]
            [Switch]$ImportCSV
      
    )

    Begin {

        if ($Global:FunctionRun -eq $null) {
            
            $errvar = $null

            import-module msonline -ErrorAction SilentlyContinue -ErrorVariable errvar

            if ($errvar) {
                Write-Host "Required modules: " -f cyan -nonewline; Write-Host "'MSOnline' " -f yellow -nonewline; Write-Host "not detected, installing." -f cyan;
                Start-Sleep -Seconds 5
                Install-Module MSOnline -ErrorAction Stop
                Import-Module MsOnline -ErrorAction Stop
                write-host "Installation complete, starting Sign-In process..." -f cyan
                Start-Sleep -Seconds 3
            }
            else {
                Write-Host "Required modules have been loaded. starting Sign-In process..." -f Cyan
                Start-Sleep -seconds 3
            }


            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
            $credential = Get-Credential
            Connect-MsolService -Credential $credential -ErrorAction Stop
            $Global:FunctionRun = $True
        }

        else {
        write-output "User currently signed in, aborting signup"
        }

    }

    Process {

        if($ImportCSV) {

            $users = Import-Csv -Path $FilePath
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
                    Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -RemoveLicenses $($SelectedSubRm.accountskuid -ErrorAction Inquire

                    write-host "Added $($SelectedSubAdd.accountskuid)"
                    Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -AddLicenses $($SelectedSubAdd.accountskuid) -ErrorAction Inquire
                }

                else {
                   write-host "User $($U.displayname) - $($U.userprincipalname) does not have this subscription enabled, aborting user."
                }
            }

        }
        else{

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
                    Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -RemoveLicenses $($SelectedSubRm.accountskuid) -ErrorAction Inquire

                    write-host "Added $($SelectedSubAdd.accountskuid)"
                    Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -AddLicenses $($SelectedSubAdd.accountskuid) -ErrorAction Inquire
                }

                else {
                   write-host "User $($U.displayname) - $($U.userprincipalname) does not have this subscription enabled, aborting user."
                }
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
