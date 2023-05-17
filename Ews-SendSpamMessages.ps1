<#
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
//
.Synopsis
  Send one or more messages using EWS
.DESCRIPTION
   Messages can be sent from random mailboxes in the organization or from a specified sender.
   Recipients can be random mailboxes in the organization or to a specific recipient.
   The messages can include an attachment.

.EXAMPLE
   .\Ews-SendSpamMessages.ps1 -Sender jim@contoso.com -MailboxLocation OnPremises -OAuth:$false -EWSUrl outlook.contoso.com -UseImpersonation:$false
.EXAMPLE
   .\Ews-SendSpamMessages.ps1 -Sender jim@contoso.com -MailboxLocation Cloud -UseImpersonation:$false -ApplicationPermission Delegated -Recipient jeff@contoso.com
.EXAMPLE
   .\Ews-SendSpamMessages.ps1 -Sender jim@contoso.com -MailboxLocation Cloud -ApplicationPermission Application -IncludeAttachments:$true -AttachmentPath C:\Scripts\Attachments\

.INPUTS
   Sending any message
   Required Parameter   -NumberOfMessages
   Required Parameter   -MailboxLocation
   Optional Parameter   -Sender
   Optional Parameter   -Recipient
   Optional Parameter   -IncludeAttachments
   Optional Parameter   -AttachmentPath
   Optional Parameter   -UseImpersonation
   Optional Parameter   -EnableLogging
   Optional Parameter   -OutputPath
   Optional Parameter   -OAuth
   Optional Parameter   -ApplicationPermission
   
   Sending from an on-premises mailbox
   Required Parameter   -EwsURL
   Optional Parameter   -Credential

.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Version: 20230517.0910
#>

param(
    [Parameter(Mandatory = $false)] [System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory = $false, HelpMessage="Sender to search for in items")] [string] $Sender,
    [Parameter(Mandatory = $false, HelpMessage="Sender to search for in items")] [string] $Recipient,
    [Parameter(Mandatory = $false, HelpMessage="Account used has impersonation rights")] [boolean] $UseImpersonation=$true,
    [Parameter(Mandatory = $false, HelpMessage="Enables EWS trace logging")] [boolean] $EnableLogging=$false,
    [Parameter(Mandatory = $true, HelpMessage="Location of the mailbox")] [ValidateSet("OnPremises", "Cloud")] [string]$MailboxLocation="Cloud",
    #[Parameter(Mandatory = $false, HelpMessage="Use OAuth for authentication")] [boolean] $OAuth= $(if($MailboxLocation -eq "Cloud") {$true} else {$false}),
    [Parameter(Mandatory = $false, HelpMessage="Use OAuth for authentication")] [boolean] $OAuth= $true,
    [Parameter(Mandatory = $false, HelpMessage="EWS namespace for on-premises environment (ex: ews.contoso.com)")] [string] $EwsURL = $(if($MailboxLocation -eq "Cloud"){"outlook.office365.com"} else {throw "-EwsURL must be passed for on-premises mailbox."}),
    [Parameter(Mandatory = $false, HelpMessage="Application permission type of either Delegated or Application")] [ValidateSet("Delegated", "Application")] [String]$ApplicationPermission="Application",
    [Parameter(Mandatory = $false, HelpMessage="Location for the log file")] [String]$OutputPath,
    [Parameter(Mandatory = $false, HelpMessage="Enables the script to send attachments with the messages")] [boolean] $IncludeAttachments=$false,
    [Parameter(Mandatory = $false, HelpMessage="Location where attachments are stored")] [String]$AttachmentPath="C:\Scripts\Attachments\",
    [Parameter(Mandatory=$true)] [int] $NumberOfMessages=3
)

function TraceHandler(){
$sourceCode = @"
    public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        public System.String LogFile {get;set;}
        public void Trace(System.String traceType, System.String traceMessage)
        {
            System.IO.File.AppendAllText(this.LogFile, traceMessage);
        }
    }
"@    

    Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $ewsDLL #$Script:EWSDLL
    $TraceListener = New-Object ewsTraceListener
   return $TraceListener
}
function CreateWord {
Param(
 [Parameter(Mandatory=$true)] [int]$LetterCount
)
	$Word = -join ((65..90) + (97..122) | Get-Random -Count $LetterCount | % {[char]$_})
	return $Word
}
function Get-ApplicationOAuthToken {
    #Change the AppId, AppSecret, and TenantId to match your registered application
    $AppId = "2f79178b-54c3-4e81-83a0-a7d16010a424"
    $AppSecret = ".hU8Q~wCUScLFX.~SqnpOHf3e~_ijjGgjPO9YcNz"
    $TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
    $Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $Body = @{
        client_id     = $AppId
        scope         = $Scope
        client_secret = $AppSecret
        grant_type    = "client_credentials"
    }
    $TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
    #Unpack the access token
    $Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
    return $Token
}
function Get-DelegatedOAuthToken {
    #Check and install Microsoft Authentication Library module
            if(!(Get-Module -Name MSAL.PS -ListAvailable -ErrorAction Ignore)){
                try { 
                    Install-Module -Name MSAL.PS -Repository PSGallery -Force
                }
                catch {
                    Write-Warning "Failed to install the Microsoft Authentication Library module."
                    exit
                }
                try {
                    Import-Module -Name MSAL.PS
                }
                catch {
                    Write-Warning "Failed to import the Microsoft Authentication Library module."
                }
            }
            $AppId = "2f79178b-54c3-4e81-83a0-a7d16010a424"
            $RedirectUri = "msal2f79178b-54c3-4e81-83a0-a7d16010a424://auth"
            $Token = Get-MsalToken -ClientId $AppId -RedirectUri $RedirectUri -Scopes $Scope -Interactive
            return $Token.AccessToken
        
}

$Scope = "https://$EwsURL/.default"
if($OAuth) {
    switch ($ApplicationPermission) {
        "Application" { $Global:OAuthToken = Get-ApplicationOAuthToken }
        "Delegated" { $Global:OAuthToken = Get-DelegatedOAuthToken }       
    }
}
elseif ($Credential -like $null) {
    $Credential = Get-Credential -Message "Please enter credentials to access mailbox(es)."
}

#region LoadEwsManagedAPI
$ewsDLL = (($(Get-ItemProperty -ErrorAction Ignore -Path Registry::$(Get-ChildItem -ErrorAction Ignore -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
if ($ewsDLL -notlike $null) {
    $ewsDLL = $ewsDLL + "Microsoft.Exchange.WebServices.dll"
    Import-Module $ewsDLL
}
else {
    if(Get-Item .\Microsoft.Exchange.WebServices.dll -ErrorAction Ignore) {
        $ewsDLL = "$(Get-Location)\Microsoft.Exchange.WebServices.dll"
         Import-Module $ewsDLL
    }
    else {
        Write-Warning "This script requires the EWS Managed API 1.2 or later."
        exit
    }
}
#endregion

#region GetSenderAndRecipients
if($Sender -like $null -or $Recipient -like $null) {
    if(!(Get-Command Get-Mailbox -ErrorAction Ignore)) {
        if($MailboxLocation -eq "Cloud") {
            try {Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -PSSessionOption $SessionOptions -ShowBanner:$False}
            catch {Write-Warning "Failed to connect to Exchange Online. Please connect and try again."; exit }
        }
        elseif ($MailboxLocation -eq "OnPremises") {
            try { Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$EwsURL/PowerShell -AllowRedirection -Authentication Kerberos) }
            catch {Write-Warning "Failed to connect to Exchange On-Premises. Please connect and try again."; exit}
        }
    }
    if($Sender -like $null) {
        Write-Host "Getting list of available mailboxes to send as..." -NoNewline -ForegroundColor Cyan
        if($Global:Mailboxes -like $null) {
            $Global:Mailboxes = (Get-Mailbox -ResultSize unlimited | where {$_.Name -notlike "DiscoverySearch*" -and $_.Name -notlike "jmartinadmin"}).PrimarySmtpAddress
        }
        Write-Host "COMPLETE" -ForegroundColor Green
    }
    if($Recipient -like $null) {
        Write-Host "Getting a list of recpipients to send to..." -NoNewline -ForegroundColor Cyan
        if($Global:Mailboxes -like $null) {
            $Global:Mailboxes = (Get-Mailbox -ResultSize unlimited | where {$_.Name -notlike "DiscoverySearch*" -and $_.Name -notlike "jmartinadmin"}).PrimarySmtpAddress
        }
        $Recipients = $Global:Mailboxes
        #$Recipients = (Get-Recipient -ResultSize Unlimited | Where {($_.RecipientType -eq "UserMailbox" -and $_.Name -notlike "DiscoverySearch*") -or $_.RecipientType -eq "MailUser"}).PrimarySmtpAddress
        Write-Host "COMPLETE" -ForegroundColor Green
    }
}
#endregion

## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013  

## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
$service.UserAgent = "EwsPowerShellScript"
$service.Url = "https://$EwsURL/ews/exchange.asmx"

#region Enable the EWS trace for debugging
if($EnableLogging) {
    if($OutputPath -like $null) {
        $OutputPath = Get-Location
    }
    $service.TraceEnabled = $true
    $TraceHandlerObj = TraceHandler
    $TraceHandlerObj.LogFile = "$OutputPath\$Sender.log"
    $service.TraceListener = $TraceHandlerObj
}
#endregion

if($IncludeAttachments) {
    if($AttachmentPath -like $null) {
        $AttachmentPath = Get-Location
    }
    $global:Attachments = (Get-ChildItem $AttachmentPath).FullName
}

for ($i=1;$i -le $NumberOfMessages; $i++) {
    $pc = ($i/$NumberOfMessages)*100 
    if($stopWatch.ElapsedMilliseconds -gt 300000 -and $OAuth) {
        Write-Host "Renewing the OAuth token..." -ForegroundColor Cyan -NoNewline
        $stopWatch.Stop(); 
        $Global:OAuthToken = Get-OAuthToken
        Write-Host "COMPLETE"
        $stopWatch.Restart()
    }
    if($Sender -like $null) {
        $MailboxName = Get-Random -Count 1 -InputObject $Global:Mailboxes
    }
    else {$MailboxName = $Sender}
    $service.HttpHeaders.Clear()
    if($OAuth) {
        $service.HttpHeaders.Add("Authorization", "Bearer $($Global:OAuthToken)")
    }
    else {
        $creds = New-Object System.Net.NetworkCredential($Credential.UserName.ToString(),$Credential.GetNetworkCredential().password.ToString())
        $service.Credentials = $creds
    }
    $service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName);
    if($UseImpersonation) {
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
    }
    #region Connect to the user's sent items folder
    $SentItemsId= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)
    $SentConnected = $true
    try {$SentItemsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$SentItemsId)
     }
    catch { Write-Warning "Unable to connect to the Sent Items folder for $($MailboxName)"; 
        $SentConnected = $false
     }
    #endregion
    #region SendMessage
    if($SentConnected){
        $message = $null
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
        [System.Collections.ArrayList]$Subject = New-Object System.Collections.ArrayList
        $SubjectWordCount = Get-Random -Minimum 1 -Maximum 6
        for($x=1; $x -le $SubjectWordCount; $x++) {
            $SubjectWord = CreateWord (Get-Random -Minimum 1 -Maximum 8)
            $Subject.Add($SubjectWord) | Out-Null
        }
        [string]$MessageSubject = $Subject
        $message.Subject = $MessageSubject
        [System.Collections.ArrayList]$Body = New-Object System.Collections.ArrayList
        $BodyWordCount = Get-Random -Minimum 5 -Maximum 500
        for($x=1; $x -le $BodyWordCount; $x++) {
            $BodyWord = CreateWord (Get-Random -Minimum 1 -Maximum 8)
            $Body.Add($BodyWord) | Out-Null
        }
        [string]$MessageBody = $Body
        $message.Body = $MessageBody
        if($IncludeAttachments) {
            $Attachment = Get-Random -Count 1 -InputObject $Global:Attachments
            $message.Attachments.AddFileAttachment($Attachment) | Out-Null
        }
        #$message.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML
        if($Recipient -like $null) {
            $ToRecipients = Get-Random -Count (Get-Random -Minimum 1 -Maximum 5) -InputObject $Recipients
        }
        else { $ToRecipients = $Recipient}
        foreach($r in $ToRecipients) {
            $message.ToRecipients.Add($r) | Out-Null
        }
        Write-Progress -Activity "Spamming $($Recipients)..." -CurrentOperation "$i of $NumberOfMessages complete" -Status "Sending message from $($Sender) to $($ToRecipients)" -PercentComplete $pc
        $message.SendAndSaveCopy($SentItemsFolder.Id) #| Out-Null
        Start-Sleep -Seconds (Get-Random -Minimum 1 -Maximum 10)
    }
    #endregion
}
