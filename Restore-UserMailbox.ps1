<#
.Synopsis
   LOW BUSINESS IMPACT - MAILBOX RESTORE TO SPECIFIED FOLDER
   Function Created per User Request
   Mailbox restoration
.DESCRIPTION
   This script will first provide you with mailbox GUID, then you will need to use that GUID to preform recovery action.

   
    Step 1: Get Exchange GUID
       Restore-UserMailbox -GUID -MailboxPrimarySMTPAddress John@Contoso.com

    Step 2: Use received Exchange GUID from previous step to start recovery actions
       Restore-UserMailbox -MailboxPrimarySMTPAddress John@contoso.com -MailboxExchangeGUID a1b23c45-1111-2ds1-b8c0-4d7107167763 -TargetRestoreFolder "Testrestorefolder"
.EXAMPLE

Step 1: Get GUID
   Restore-UserMailbox -GUID -MailboxPrimarySMTPAddress John@Contoso.com

Step 2: Use received GUID to start recovery actions
   Restore-UserMailbox -MailboxPrimarySMTPAddress John@contoso.com -MailboxExchangeGUID a1b23c45-1111-2ds1-b8c0-4d7107167763 -TargetRestoreFolder "Testrestorefolder"

.INPUTS
   -MailboxPrimarySMTPAddress | -MailboxExchangeGUID | -TargetRestoreFolder
.OUTPUTS
   None
.NOTES
   v-bakwi
   Please note that this script has Low Business Impact as it modifies user mailbox in non-sensitive way. Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   N/A
.ROLE
   Exchange mailbox management
.FUNCTIONALITY
   Restore mailbox contents
#>
function Restore-UserMailbox
{
    [CmdletBinding()]
    Param
    (
        
        # Use this switch to check mailbox Exchange GUID
        [Parameter(Mandatory=$true, ParameterSetName = 'GUID')]
        [Switch]$GUID,
       
        # Provide E-Mail address of the mailbox you want to restore
        [Parameter(Mandatory=$true, ParameterSetName = 'GUID')]
        [Parameter(Mandatory=$true, ParameterSetName = 'Restore')]
        $MailboxPrimarySMTPAddress,

        # Provide Exchange GUID of mailbox ypou want to restore
        [Parameter(Mandatory=$true, ParameterSetName = 'Restore')]
        $MailboxExchangeGuid,

        # Mailbox folder in "" to which you want to provision restored files
        [Parameter(Mandatory=$true, ParameterSetName = 'Restore')]
        $TargetRestoreFolder
        
        
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

        If ($GUID) {

        $script:FunctionRun = "True"
        $Mailbox = Get-Mailbox -Identity $MailboxPrimarySMTPAddress -InactiveMailboxOnly | fl Name,DistinguishedName,ExchangeGuid,PrimarySmtpAddress
        return $Mailbox

        }

        Else {

        $InactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity $MailboxExchangeGuid

        }
    }
    Process{
        $script:FunctionRun = "True"

        New-MailboxRestoreRequest -SourceMailbox $InactiveMailbox.DistinguishedName -TargetMailbox $MailboxPrimarySMTPAddress -TargetRootFolder $TargetRestoreFolder -AllowLegacyDNMismatch

        write-host "Mailbox restore request has been scheduled, please review it down below:"

        write-output Get-MailboxRestoreRequest
            
    }
    End{

        write-host "Mailbox restore request has been scheduled, please review it down below:"

        Get-MailboxRestoreRequest | write-output 

    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Restore-UserMailbox -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION"
write-host ""
get-help Restore-UserMailbox -ShowWindow
