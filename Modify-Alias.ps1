<#
.Synopsis 
   LOW BUSINESS IMPACT - MAILBOX SECONDARY ALIAS MODIFICATION
   Function Created per User Request
   Alias Removal/Addition
.DESCRIPTION
   This script will remove a specified alias for all users of one specified domain. With -Add Parameter it will add the alias instead of removing it.
.EXAMPLE
   To remove the aliases
   
   Modify-Alias -AliasDomain TestDomain.onmicrosoft.com -FilePath C:\TestLocation -InitialDomain InitialDomain.com
.EXAMPLE
   To add the aliases

   Modify-Alias -AliasDomain TestDomain.onmicrosoft.com -FilePath C:\TestLocation -InitialDomain InitialDomain.com -Add
.INPUTS
   -InitialDomain, -AliasDomain, -FilePath [SWITCH]-Add
.OUTPUTS
   Mailboxes.csv at -FilePath location
.NOTES
   v-bakwi
   Please note that this script has Low Business Impact as it modifies user mailbox in non-sensitive way. Please note that you use all scripts, provided or not, on your own responsibility.
   At least basic Powershell knowledge is required to monitor script behaviour.
.COMPONENT
   N/A
.ROLE
   Exhange Online Mailbox Management
.FUNCTIONALITY
   Add/Remove Additional Alias
#>
function Modify-Alias {
    [CmdletBinding()]
    Param
    (
        #Input the domain to be removed/added
        [Parameter(Mandatory=$true)]
        $AliasDomain,

        #Specify the filepath for Temporary work files to be saved in
        [Parameter(Mandatory=$true)]
        $FilePath,

        #Input the primary domain on which users have their Primary SMTP Addresses on mailbox
        [Parameter(Mandatory=$true)]
        $InitialDomain,

        #Select to add aliases instead of removing
        [Parameter()][Switch]
        $Add
        
    )

    Begin{
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
        Set-Location -Path $Filepath
        
        $Cred = Get-Credential
        
        Install-Module MsOnline -ErrorAction Stop
        Import-Module MSOnline -ErrorAction Stop
        Connect-MsolService –Credential $Cred -ErrorAction Stop 
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred –AllowRedirection -ErrorAction Stop
        Import-PSSession $ExchangeSession -ErrorAction Stop 

        write-host "Exported mailbox database file to "$FilePath\Mailboxes.csv""
        Get-Mailbox | Where {$_.PrimarySmtpAddress -like "*$InitialDomain"} |Select name,primarysmtpaddress | Export-Csv "$FilePath\Mailboxes.csv" -NoTypeInformation
        $Mailbox = Import-Csv "$FilePath\Mailboxes.csv"
       
    }
    
    Process{

        If ($Add) {
            write-host "Add Mailbox Alias Operation Started"                        
            ForEach ($M in $Mailbox) {
                $UserName =($($M.PrimarySmtpAddress) -split "@")[0]    
                $UPN= $UserName+"@"+$AliasDomain
                set-mailbox -identity $($M.PrimarySmtpAddress) -EmailAddresses @{add = $UPN} -ErrorAction Inquire
                write-host "Adding alias $UPN to mailbox $($M.Name)"
            }
        }
        Else {
            write-host "Remove Mailbox Alias Operation Started"
            ForEach ($M in $Mailbox) {
                $UserName =($($M.PrimarySmtpAddress) -split "@")[0]    
                $UPN= $UserName+"@"+$AliasDomain
                set-mailbox -identity $($M.PrimarySmtpAddress) -EmailAddresses @{remove = $UPN} -ErrorAction Inquire
                write-host "Removing alias $UPN from mailbox $($M.name)"
            }

        }
    }
    End{
        
        write-host "All mailboxes on domain $InitialDomain have been modified with alias based on $AliasDomain domain."
        write-host ""
        write-host "You can now safely close the console"
    }
}
Clear-Host
Write-Host ""
write-host "Please execute Modify-Alias -AliasDomain TestDomain.onmicrosoft.com -FilePath C:\TestLocation -InitialDomain Migratedcomain.com (-Add)"
Write-Host ""
Get-Help Modify-Alias -ShowWindow
