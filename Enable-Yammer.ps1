<#
.Synopsis
   HIGH BUSINESS IMPACT - USER LICENSE/SERVICE MODIFICATION
   Function Created per User Request
   Team Visibility Modification
.DESCRIPTION
   This script will view and edit all user sku's to enable Yammer
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
   Edit user SKU to allow usage of Yammer
#>
function Enable-Yammer
{
    [CmdletBinding()]
    Param ()

    Begin {

        if ($Global:FunctionRun -eq $null) {
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
        
        foreach ($user in $users){

                Write-Host "Checking " $user.UserPrincipalName -foregroundcolor "Cyan"
                $CurrentSku = $user.Licenses.Accountskuid


                Write-Host $user.UserPrincipalName "Has more than one license assigned. Looping through all of them." -foregroundcolor "White"
                for($i = 0; $i -lt $currentSku.count; $i++){

                                
                                $CurrentServicesInCurrentSKU = $null
                                $CurrentServicesInCurrentSKU = $user.Licenses[$i].ServiceStatus.ServicePlan.ServiceName

                                if ($CurrentServicesInCurrentSKU -notlike "*YAMMER_ENTERPRISE*"){
                                    
                                        #Pulls up the data for service lict and filters only disabled and only these that do not have Yammer in name - creates list of disabled SKU's for user
                                        $DisabledServices =  $user.Licenses.servicestatus | where {$_.serviceplan.servicename -notlike "*YAMMER*" -and $_.provisioningstatus -like "*Disabled*" }

                                        #Inputs $DisabledServices as -DisabledPlans Array for new MsolLicensePlan for user
                                        $NewSkU = New-MsolLicenseOptions -AccountSkuId $user.Licenses[$i].AccountSkuid -DisabledPlans $DisabledServices.ServicePlan.servicename
                                        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $NewSkU  

                                }

                                else {
                                 write-host "Yammer currently enabled, aborting"
                                }
        
                       }

        }
        
    }

    End
    {
    write-host "All User SKU's have been reviewed"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Enable-Yammer -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION"
write-host ""
get-help Enable-Yammer -ShowWindow


