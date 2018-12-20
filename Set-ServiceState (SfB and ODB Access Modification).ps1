<#
.Synopsis 
   HIGH BUSINESS IMPACT - SERVICE ACCESSIBILITY MODIFICATION
   Function Created per User Request
   Removal/Reassign of SfB User Licenses and ODB Access
.DESCRIPTION
   The function will firstly remove SfB feature from every user, then get their Personal Site and change the ownership, rendering them unable to access both SfB and ODB.
   Next, when running -Revert, the function will get the info back from SPOSites.csv file generated on first run and retrieve original site owners, restoring access to OSB Sites. It will also reassign SfB licenses, enabling the service.
.EXAMPLE
   Set-ServiceState -NewOwner testaccount@testdomain.com -OrganizationName TestCompanyName -FilePath C:\TestFolder -Revert
.EXAMPLE
   Set-ServiceState -NewOwner testaccount@testdomain.com -OrganizationName TestCompanyName -FilePath C:\Users\Admin\Desktop\ScriptResultFolder 
.INPUTS
   -NewOwner, -OrganizationName, -FilePath
.OUTPUTS
   ExportedUsers.csv & SPOSites.csv at -FilePath location
.NOTES
   v-bakwi
   Please note that this script has High Business Impact as it modifies service accessibility. Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor function behaviour.
.COMPONENT
   N/A
.ROLE
   Office 365 Service Administration
.FUNCTIONALITY
   Service Accessibility Modification
#>
Function Set-ServiceState {

    [CmdletBinding()]

    Param (
        #Will revert made changes if param provisioned
        [Parameter()][Switch]$Revert,

        [Parameter(Mandatory=$True)]
        #Provide new owner for ODB Sites in format >test@domain.com<
        [String]$NewOwner,
        
        [Parameter(Mandatory=$True)]
        #Input your company's Sharepoint company name (companyname.sharepoint[...] in the URL)
        [String]$OrganizationName,
       
        [Parameter(Mandatory=$True)]
        #Specify FilePath for saving User and Site CSV Files (Please make sure that when running -Revert, FilePath is exactly the same as it was for 1st time)
        [String]$FilePath
    )

    Begin{  

        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
        Set-Location -Path $FilePath
        
        $Cred = Get-Credential
        $adminUPN = $Cred.UserName
        $AccountSKU = ("$OrganizationName"+":O365_BUSINESS_PREMIUM")

        Install-Module MsOnline
        Import-Module MSOnline
        Install-Module Microsoft.Online.SharePoint.PowerShell
        Import-Module Microsoft.Online.SharePoint.PowerShell
        Connect-MsolService –Credential $Cred 
        Connect-SPOService -Url https://$OrganizationName-admin.sharepoint.com -Credential $Cred
        
        Write-Host "Extracting all users from Tenant, exporting them to a ExportedUsers.csv"
        Get-MSOlUser -All | Where {$_.UserPrincipalName -notlike "*#EXT#*"} | Where {$_.isLicensed -eq $True} | Select-Object UserPrincipalName | Export-Csv -Path ".\ExportedUsers.csv" -NoTypeInformation -Force
        
    }

    Process{
        If($Revert){

            write-host "Operation revert selected"

            $User = Import-Csv -Path ".\ExportedUsers.csv"
            $SPOSite = Import-CSV -Path ".\SPOSites.csv"

            write-host "User.csv and SPOSites.csv imported"

            $LO = New-MsolLicenseOptions -AccountSkuId $AccountSKU -DisabledPlans $null

            ForEach ($S in $SPOSite) {

                Set-SPOSite -Identity $($S.Url) -Owner $($S.Owner)

                write-host "Site $($S.Url) ownership set to $($S.Owner))" 

            }

            ForEach ($U in $User) {
                
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -LicenseOptions $LO

                write-host "User $U has their Skype for Business License active"

            }
        }
        Else {
            write-host "Operation started, creating SPOSites.csv file"

            Get-SPOSite -IncludePersonalSite:$true -Limit ALL |Where {$_.Url -like "*personal*"} | select url, Owner | Export-Csv -Path “.\SPOSites.csv” -NoTypeInformation -Force
            $User = Import-Csv -Path ".\ExportedUsers.csv"
            $SPOSite = Import-CSV -Path ".\SPOSites.csv"

            write-host "User.csv and SPOSites.csv imported"

            $LO = New-MsolLicenseOptions -AccountSkuId $AccountSKU -DisabledPlans "MCOSTANDARD"
            ForEach ($S in $SPOSite) {

                Set-SPOSite -Identity $($S.Url) -Owner $NewOwner
                write-host "Site $($S.Url) ownership set to $NewOwner)" 

            }

            ForEach ($U in $User) {
                
                Set-MsolUserLicense -UserPrincipalName $($U.userprincipalname) -LicenseOptions $LO

                write-host "User $U has their Skype for Business License deactivated"

            }

        }
    }

    End{
        If($Revert) {
            write-host "Operation revert has been completed successfully, please confirm with users that they have regained access to SfB & ODB"
        }
        Else {
            write-host "Operation has been completed successfully, please confirm with users that they have lost access to SfB & ODB"
        }
    }
    
}
Clear-Host
Write-Host ""
write-host "Please execute Set-ServiceState -NewOwner testaccount@testdomain.com -OrganizationName TestCompanyName (-Revert)"
Write-Host ""
Get-Help set-servicestate -ShowWindow
