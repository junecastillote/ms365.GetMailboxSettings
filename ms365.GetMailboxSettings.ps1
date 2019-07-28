
<#PSScriptInfo

.VERSION 1.0

.GUID e465e748-b3f7-46bb-b0c8-39d945c4b26b

.AUTHOR June Castillote

.COMPANYNAME lazyexchangeadmin.com

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/ms365.GetMailboxSettings

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>

<# 

.DESCRIPTION
    Retrieve Office 365 Mailbox Settings using MS Graph API calls. Recommended if there's a large number of users to process.
    This can usually be achieved using the Exchange Online powershell session/commands. However, it is quite unreliable when
    when processing a large number of users - due to throtlling, data size limitation and session timeouts to name a few.

    * This function utilize the MS Graph API calls which all goes through HTTPS.
    * Automatically renews the access token every 58 minutes to ensure that authorization is current and not expired.
    * Returns the final result as an object which can be manipulated, filtered, or exported to different formats as required.
.SYNOPSIS
    Retrieve Office 365 Mailbox Settings using MS Graph API calls
.EXAMPLE
    PS C:\> $mailboxSettings = .\ms365.GetMailboxSettings.ps1 -ClientID <clientID> -ClientSecret <clientSecret> -tenantID <tenantID> -All
    This example will authenticate with MS Graph API and retrieve ALL users' mailbox settings.
.EXAMPLE
    PS C:\> $mailboxSettings = .\ms365.GetMailboxSettings.ps1 -ClientID <clientID> -ClientSecret <clientSecret> -tenantID <tenantID> -UserID "user1@domain.com","f37ff902-28f3-4195-8c59-ce8afacbfd35"
    This example will authenticate with MS Graph API and retrieve the mailbox settings of the two users.
    User 1 with UserPrincipalName of "user1@domain.com"
    User 2 with ObjectGUID of "f37ff902-28f3-4195-8c59-ce8afacbfd35"
.INPUTS
    None
.OUTPUTS
    None
.NOTES
    None
#> 

[cmdletbinding()]
param (
    [parameter(mandatory = $true)]
    [string]$ClientID, #The Application ID of the registered azure ad app
    
    [parameter(mandatory = $true)]
    [string]$ClientSecret, #The Secret Key of the registered azure ad app

    [parameter(mandatory = $true)]
    [string]$TenantID, #TenantID - the Directory ID or domain of the tenant.

    [Parameter()]
    [string[]]$userID, #Use userID to specify a user or an array of users to process.
    #Cannot be used together with the 'All' switch
    [Parameter()]
    [switch]$All, #This is the default if userID is not specified.
    #This will get All users from Azure AD with usertype 'member'

    [Parameter()]
    [int]$MaxPage   #This puts a limit to how many pages of users will be retrieved by the All switch.
    #Each page contains 100 records.
    #If not specified, All pages will be retrieved.
)
#...................................................................................
#Region Function
Function Get-oAuth {	
    param(
        [parameter(mandatory = $true)]
        [string]$ClientID,
        [parameter(mandatory = $true)]
        [string]$ClientSecret,
        [parameter(mandatory = $true)]
        [string]$TenantID
    )
    
    try {
        $body = @{grant_type = "client_credentials"; scope = "https://graph.microsoft.com/.default"; client_id = $ClientID; client_secret = $ClientSecret }
        $oAuth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Body $body
        
        #set expire time to 2 minutes earlier
        $expireDateTime = ((Get-Date).AddSeconds($oAuth.expires_in)).AddSeconds(-120)
        #compose token
        $token = @{'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }    
        #add expireDateTime property
        $oAuth | Add-Member -Name expireDateTime -MemberType NoteProperty -Value $expireDateTime    
        #add token property
        $oAuth | Add-Member -Name token -MemberType NoteProperty -Value $token
        return $oAuth
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        return $null
    }
}
#EndRegion Function
#...................................................................................

#...................................................................................
#Region Parameter check
#do not allow 'userID' and 'All' in the same set.
if ($userID -and $All) {
    Write-Host "userID and All cannot be used in the same set." -ForegroundColor Yellow
    return $null
}

#do not allow 'userID' and 'Max' in the same set.
if ($userID -and $MaxPage) {
    Write-Host "userID and Max cannot be used in the same set." -ForegroundColor Yellow
    return $null
}

#if userID and All are not specified, default to All.
if (!$userID -and !$All) {
    $All = $true
}
#EndRegion Parameter check
#...................................................................................


#Get new token
if (!($oAuth = get-oAuth -ClientID $ClientID -ClientSecret $ClientSecret -TenantID $TenantID)) {
    Write-Host "There was an error getting authorization."
    Return $null
}

#...................................................................................
#Region Get All Users
if ($All) {
    try {		
        if ((Get-Date) -gt ($oauth.expireDateTime)) {
            #Get new token
            Write-Host "Renew token" -ForegroundColor Yellow
            if (!($oAuth = get-oAuth -ClientID $ClientID -ClientSecret $ClientSecret -TenantID $TenantID)) {
                Write-Host "There was an error getting authorization."
                Return $null
            }
        }
        $userID = @()
        $request = 'https://graph.microsoft.com/beta/users?$filter=usertype eq ''member''&$select=userPrincipalName,mail'
        $result = Invoke-RestMethod -Method Get -Uri $request -Headers $oAuth.Token
        $userID += ($result.value | Where-Object { $_.mail -ne $null }).UserPrincipalName
        $nextLink = $result."@odata.nextLink"
        $page = 1

        Write-Progress -Activity "Getting users..." -Status "Page $page" -PercentComplete 100
        
        #cycle through all pages
        
        #While ($nextLink) {
        if ($MaxPage) {
            While ($page -ne $MaxPage) {
                $page++         
                $result = Invoke-RestMethod -Method Get -Uri $nextLink -Headers $oAuth.Token
                $userID += ($result.value | Where-Object { $_.mail -ne $null }).UserPrincipalName
                $nextLink = $result."@odata.nextLink"
                Write-Progress -Activity "Getting users..." -Status "Page $page" -PercentComplete 90
            }
        }
        else {
            While ($nextLink) {
                $page++         
                $result = Invoke-RestMethod -Method Get -Uri $nextLink -Headers $oAuth.Token
                $userID += ($result.value | Where-Object { $_.mail -ne $null }).UserPrincipalName
                $nextLink = $result."@odata.nextLink"            
                Write-Progress -Activity "Getting users..." -Status "Page $page" -PercentComplete 90
            }
        }        
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        return $null
    }
}
#EndRegion Get All Users
#...................................................................................

#...................................................................................
#Region Get Mailbox Settings
#store count of userId. These Ids may or may not have a mailbox.
$userCount = ([array]$userID).Count
$mailboxSettings = @()
$index = 1
foreach ($id in $userID) {
	
    if ((Get-Date) -gt ($oauth.expireDateTime)) {
        #Get new token
        Write-Host "Renew token" -ForegroundColor Yellow
        if (!($oAuth = get-oAuth -ClientID $ClientID -ClientSecret $ClientSecret -TenantID $TenantID)) {
            Write-Host "There was an error getting authorization."
            Return $null
        }
    }
	
    $percentComplete = [int]($index / $userCount * 100)
    Write-Progress -Activity "Processing..." -Status "($index of $userCount [$percentComplete%]) - $id)" -PercentComplete ($index / $userCount * 100)	
    $request = "https://graph.microsoft.com/beta/users/$id/mailboxSettings"
    try {
        $settings = Invoke-RestMethod -Method Get -Uri $request -Headers $oAuth.Token
        $settings | Add-Member -Name UserPrincipalName -MemberType NoteProperty -Value $id
        $mailboxSettings += $settings
    }
    catch {
        Write-Host "$id - $($_.exception.Message)" -ForegroundColor Yellow
    }
    $index++
}
#EndRegion Get Mailbox Settings
#...................................................................................
Write-Progress -Activity "Completed." -Status "Done" -PercentComplete 100 -Completed
return $mailboxSettings


