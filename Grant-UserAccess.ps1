<#
.Synopsis
   LOW BUSINESS IMPACT - MAILBOX FOLDER ACCESS GRANT
   Function Created per User Request
   Mailbox restoration
.DESCRIPTION
   This script, based on input, will give appropriate access rights to all folders under single mailbox.

   Grant-UserAccess -MailboxOwner John@Contoso.com -AddedUser Iwanttheaccess@contoso.com -Folderfilter * -AccessLevel Reviewer -FileDir "C:\Users\Administrator\Desktop\Script Test"
    
.EXAMPLE

Grant-UserAccess -MailboxOwner John@Contoso.com -AddedUser Iwanttheaccess@contoso.com -Folderfilter * -AccessLevel Reviewer -FileDir "C:\Users\Administrator\Desktop\Script Test"

.INPUTS
   -MailboxOwner | -AddedUser | -Folderfilter | -AccessLevel | -FileDir
.OUTPUTS
   Folders.csv file at -FileDir location
.NOTES
   v-bakwi
   Please note that this script has Low Business Impact as it modifies user mailbox in non-sensitive way. Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   N/A
.ROLE
   Exchange mailbox management
.FUNCTIONALITY
   Grant access to mailbox folders
#>
function Grant-UserAccess
{
    [CmdletBinding()]
    Param
    (
        
        # E-Mail address of mailbox owner
        [Parameter(Mandatory=$true)]
        $MailboxOwner,
       
        # E-Mail address of user that will receive access
        [Parameter(Mandatory=$true)]
        $AddedUser,

        # Folder filters, ie. Inbox, Recycle Bin etc. Type * if you want to receive access to all folders.
        [Parameter(Mandatory=$true)]
        $FolderFilter,

        # File directory in "" to which export file will be saved for further use as script database
        [Parameter(Mandatory=$true)]
        $FileDir,

        # Define what permissions would you like to provide: Reviewer, Author, Editor, Contributor, Owner etc. (https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxfolderpermission?view=exchange-ps)
        [Parameter(Mandatory=$true)]
        $AccessLevel
        
        
    )

    Begin{
        clear-host

        If ($script:FunctionRun -eq $null) {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
       
            $credential = get-credential 
            Install-Module MSOnline -ErrorAction Stop
            Import-Module MsOnline -ErrorAction Stop
            Connect-MsolService -Credential $credential -ErrorAction Stop   
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" –AllowRedirection -ErrorAction Stop  
            Import-PSSession $ExchangeSession -ErrorAction Stop
        }

        Else {}

        Write-host "Dropping all found folders to CSV file"
        Get-MailboxFolderStatistics -Identity $MailboxOwner -FolderScope All | Select-Object folderpath | where {$_.folderpath -like "$FolderFilter"} | % {$_.folderpath.replace('/','\')} |Out-File "$FileDir\Folders.csv" -Encoding UTF8
        Write-host "Importing database back to powershell"
        $folderarray = Get-Content "$FileDir\Folders.csv"

    }
    Process{
        $script:FunctionRun = "True"

        write-host "Adding $AccessLevel permissions to folders located on $MailboxOwner mailbox for user $AddedUser"
        
        ForEach ($folder in $folderarray) {
            write-host ""
            write-host "Working on $folder folder"
            Add-MailboxFolderPermission -Identity ($($MailboxOwner)+":"+$($folder)) -User $AddedUser -AccessRights $AccessLevel -Confirm:$false -ErrorAction Inquire
            write-host ""
        }

    }
    End{
        write-host "All folders for mailbox $MailboxOwner have been configured with Access Rights on access level: $AccessLevel for User $AddedUser"
        write-host "To check if function was successful, please use this cmdlet >Get-MailboxFolderPermission -Identity >Mailbox SMTP Address< | fl User,AccessRights,FolderName<"
    }
}
cls
write-host ""
write-host "PLEASE READ THE EXAMPLES IN HELP FILE (get-help Grant-UserAccess -ShowWindow) TO UNDERSTAND USAGE & PURPOSE OF THIS FUNCTION"
write-host ""
get-help Grant-UserAccess -ShowWindow
