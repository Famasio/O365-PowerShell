<#
.Synopsis
   NO BUSINESS IMPACT - UNIFIED GROUP/TEAM OUTLOOK VISIBILITY MODIFICATION
   Function Created per User Request
   Team Visibility Modification
.DESCRIPTION
   This script will make All/Specified in CSV file/Single Team visible/hidden in Outlook.

   
Modify All Teams on Tenant
   Modify-TeamsVisibility -All -Show/-Hide -FileDir "C:\Users\Administrator\Desktop\My Script Repository"

Export all teams on tenant into CSV file to exclude some manually
   Modify-TeamsVisibility -ExportData -FileDir "C:\Users\Administrator\Desktop\MyDataBase"

Import modified CSV file and modify contained Teams
   Modify-TeamsVisibility -ImportData -Show/-Hide -FileDir "Same filedir as when performing -ExportData Action"

Modify single Team of which you know its E-Mail address
   Modify-TeamsVisibility -ModifySingle -Show/-Hide -TeamSMTPAddress "MyTeam@MyDomain.com"

.EXAMPLE

Modify All Teams on Tenant
   Modify-TeamsVisibility -All -Show/-Hide -FileDir "C:\Users\Administrator\Desktop\My Script Repository"

Export all teams on tenant into CSV file to exclude some manually
   Modify-TeamsVisibility -ExportData -FileDir "C:\Users\Administrator\Desktop\MyDataBase"

Import modified CSV file and modify contained Teams
   Modify-TeamsVisibility -ImportData -Show/-Hide -FileDir "Same filedir as when performing -ExportData Action"

Modify single Team of which you know it's E-Mail address
   Modify-TeamsVisibility -ModifySingle -Show/-Hide -TeamSMTPAddress "MyTeam@MyDomain.com"

.INPUTS
   -Show/-Hide | -All/-ModifySingle | -ExportData + -ImportData | -TeamSMTPAddress
.OUTPUTS
   TeamsArray.csv at -FileDir location.
.NOTES
   v-bakwi
   Please note that this script has Low Business Impact as it modifies user mailbox in non-sensitive way. Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   N/A
.ROLE
   Unified Group Management
.FUNCTIONALITY
   Modify Unified Group Attribute "HiddenFromExchangeClientsEnabled"
#>
function Modify-TeamsVisibility
{
    [CmdletBinding()]
    Param
    (
        # Use to export only CSV file without taking any actions
        [Parameter(Mandatory=$true, ParameterSetName = 'Export')]
        [Switch]$ExportData,

        # Use to run script with imported CSV file, that you can manually edit
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportHide')]
        [Switch]$ImportData,

        # Use to modify single Unified Group/Team, changes required parameters
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleHide')]
        [Switch]$ModifySingle,

        # Use to change all Teams on tenant to desired state (Show/Hide)
        [Parameter(Mandatory=$true, ParameterSetName = 'AllShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'AllHide')]
        [Switch]$All,

        # Use to make selected Team(s) visible in Outlook
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'AllShow')]
        [Switch]$Show,

        # Use to make selected Team(s) hidden in Outlook
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportHide')]
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleHide')]
        [Parameter(Mandatory=$true, ParameterSetName = 'AllHide')]
        [Switch]$Hide,

        # Specify the directory in "" in which Export and Temp files will be saved for usage as database 
        [Parameter(Mandatory=$true, ParameterSetName = 'Export')]
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'ImportHide')]
        [Parameter(Mandatory=$true, ParameterSetName = 'AllShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'AllHide')]
        $FileDir,

        # Provide E-Mail address in "" of Unified Group/Team you want to edit as singular unit
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleShow')]
        [Parameter(Mandatory=$true, ParameterSetName = 'SingleHide')]
        $TeamSMTPAddress
        
        

    )

    Begin{
        clear-host

        If ($script:FunctionRun -eq $null) {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
       
            $credential = get-credential 
            Install-Module MSOnline
            Import-Module MsOnline 
            Connect-MsolService -Credential $credential   
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" –AllowRedirection  
            Import-PSSession $ExchangeSession
        }

        Else {}

    }
    Process{

        

        If ($ExportData) {
            $script:FunctionRun = "True"
            Get-UnifiedGroup -ResultSize Unlimited | Select DisplayName, PrimarySmtpAddress, HiddenFromExchangeClientsEnabled | Export-csv -Path $FileDir\TeamsArray.csv -NoTypeInformation
            Write-Host "TeamsArray File Generated in $FileDir\TeamsArray.csv, please run the function again with -ImportData param to run this file as database for function"
            return
        }
        
        If ($ImportData) {
            $script:FunctionRun = "True"
            $Team = Import-Csv -Path $FileDir\TeamsArray.csv
            Write-Host "Data properly imported from TeamsArray.csv file, $($Team.Count) Teams to be processed"
        }

        If ($All) {
            $script:FunctionRun = "True"
            Write-Host "HiddenTeamsArray File Generated in the set place for further usage as database for script"
            Get-UnifiedGroup -ResultSize Unlimited | Select DisplayName, PrimarySmtpAddress, HiddenFromExchangeClientsEnabled | Export-csv -Path $FileDir\TeamsArray.csv -NoTypeInformation
            $Team = Import-Csv -Path $FileDir\TeamsArray.csv
            Write-Host "Data properly loaded from TeamsArray.csv file, $($Team.Count) Teams to be processed"
        }
    }
    End{

        If ($Show) {

            ForEach ($T in $Team) {

                Write-Host "Processing $($T.DisplayName)"

                If ($T.HiddenFromExchangeClientsEnabled -eq $False) { 
                    Write-Host "Team $($T.DisplayName) is visible in Exchange and will not be processed"
                    Write-Host ""
                }

                Else { 
                    Set-UnifiedGroup -Identity $T.PrimarySmtpAddress -HiddenFromExchangeClientsEnabled:$False
                    Write-Host "Showing $($T.DisplayName) Team in Exchange"
                    Write-Host ""
                }

            }

            Write-Host "All loaded Teams are now visible in Exchange"
        }

        If ($Hide) {

            ForEach ($T in $Team) {

                Write-Host "Processing $($T.DisplayName)"

                If ($T.HiddenFromExchangeClientsEnabled -eq $True) { 
                    Write-Host "Team $($T.DisplayName) is hidden in Exchange and will not be processed"
                    Write-Host ""
                }

                Else { 
                    Set-UnifiedGroup -Identity $T.PrimarySmtpAddress -HiddenFromExchangeClientsEnabled:$True
                    Write-Host "Hiding $($T.DisplayName) Team in Exchange"
                    Write-Host ""
                }

            }

            Write-Host "All loaded Teams are now hidden in Exchange"
        }

        If ($ModifySingle) {

            $script:FunctionRun = "True"

            If ($Show){
                
                Write-Host "Processing $TeamSMTPAddress"
                Set-UnifiedGroup -Identity $TeamSMTPAddress -HiddenFromExchangeClientsEnabled:$False
                Write-Host ""
                Write-Host "$TeamSMTPAddress visible in Exchange"

            }

            If ($Hide){
                
                Write-Host "Processing $TeamSMTPAddress"
                Set-UnifiedGroup -Identity $TeamSMTPAddress -HiddenFromExchangeClientsEnabled:$True
                Write-Host ""
                Write-Host "$TeamSMTPAddress hidden in Exchange"

            }
           
        }
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Modify-TeamsVisibility -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION"
write-host ""
get-help Modify-TeamsVisibility -ShowWindow
