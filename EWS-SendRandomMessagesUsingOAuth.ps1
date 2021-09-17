<##################################################################################
#
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
#
#################################################################################
#>

param(
    [Parameter(Mandatory=$false)] [int] $NumberOfMessages=3
    )
##Create an EWS trace
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
function Get-OAuthToken{
    #Change the AppId, AppSecret, and TenantId to match your registered application
$AppId = "6a93c8c4-9cf6-4efe-a8ab-9eb178b8dff4"
$AppSecret = "B5T7Q~PnjfVyVgSaJb73gFrElsv3STOt.FhA9"
$TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
#Build the URI for the token request
$Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$Body = @{
    client_id     = $AppId
    scope         = "https://outlook.office365.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}
$TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
#Unpack the access token
$Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
return $Token
}
function CreateWord {
Param(
 [Parameter(Mandatory=$true)] [int]$LetterCount
)
	$Word = -join ((65..90) + (97..122) | Get-Random -Count $LetterCount | % {[char]$_})
	return $Word
}
## Load Managed API dll - thanks to Glen Scales for this http://gsexdev.blogspot.com/
###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$ewsDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $ewsDLL) {
    Import-Module $ewsDLL
}
else {
    "$(get-date -format yyyyMMddHHmmss):"
    "This script requires the EWS Managed API 1.2 or later."
    "Please download and install the current version of the EWS Managed API from"
    "http://go.microsoft.com/fwlink/?LinkId=255472"
    ""
    "Exiting Script."
    exit
}

$Modules = Get-Module
if ("ExchangeOnlineManagement" -notin  $Modules.Name) {
    try {Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -PSSessionOption $SessionOptions -ShowBanner:$False}
    catch {Write-Warning "Failed to connect to Exchange Online"; break }
}
Write-Host "Getting list of available mailboxes to send as..." -NoNewline -ForegroundColor Cyan
$Mailboxes = (Get-EXOMailbox -ResultSize unlimited | where {$_.Name -notlike "DiscoverySearch*"}).PrimarySmtpAddress
Write-Host "COMPLETE" -ForegroundColor Green
Write-Host "Getting a list of recpipients to send to..." -NoNewline -ForegroundColor Cyan
#$Recipients = $Mailboxes
$Recipients = (Get-Recipient -ResultSize Unlimited | Where {($_.RecipientType -eq "UserMailbox" -and $_.Name -notlike "DiscoverySearch*") -or $_.RecipientType -eq "MailUser"}).PrimarySmtpAddress
Write-Host "COMPLETE" -ForegroundColor Green

#Set a timer to monitor token age
$stopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
$stopWatch.Start()
#Get an OAuth token
$OAuthToken = Get-OAuthToken

## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  

$service.UserAgent = "EwsPowerShellScript"
$service.Url = "https://outlook.office365.com/ews/exchange.asmx"
#region Enable the EWS trace for debugging
$outputPath = "C:\Temp"
$service.TraceEnabled = $True
$TraceHandlerObj = TraceHandler
$TraceHandlerObj.LogFile = "$outputPath\$Sender.log"
$service.TraceListener = $TraceHandlerObj
#endregion

for ($i=1;$i -le $NumberOfMessages; $i++) {
    $pc = ($i/$NumberOfMessages)*100 
    Write-Progress -Activity "Spamming..." -CurrentOperation "$i of $NumberOfMessages complete" -Status "Token age is $($stopWatch.Elapsed). Please wait." -PercentComplete $pc
    if($stopWatch.ElapsedMilliseconds -gt 3540000) {
        $stopWatch.Stop(); 
        $OAuthToken = Get-OAuthToken
        $stopWatch.Restart()
    }
    $Sender = Get-Random -Count 1 -InputObject $Mailboxes
    $service.HttpHeaders.Clear()
    $service.HttpHeaders.Add("Authorization", "Bearer $($OAuthToken)")
    $service.HttpHeaders.Add("X-AnchorMailbox", $Sender);
    $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Sender)

    #region Connect to the user's sent items folder and empty it
    $SentItemsId= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$Sender)
    $SentConnected = $true
    try {$SentItemsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$SentItemsId)
     }
    catch { Write-Warning "Unable to connect to the Sent Items folder for $($Sender)"; 
        $SentConnected = $false
     }
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
    #$message.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML
    $ToRecipients = Get-Random -Count (Get-Random -Minimum 1 -Maximum 5) -InputObject $Recipients
    foreach($r in $ToRecipients) {
        $message.ToRecipients.Add($r) | Out-Null
    }
    $message.SendAndSaveCopy($SentItemsFolder.Id) #| Out-Null
    Start-Sleep -Seconds (Get-Random -Minimum 1 -Maximum 30)
    }
    #endregion
}