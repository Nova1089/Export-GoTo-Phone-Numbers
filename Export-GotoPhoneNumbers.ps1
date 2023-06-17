<#
This script exports all phone numbers from GoTo Connect.

To obtain an access token for the GoTo API, follow the steps in the getting started and authentication guides:
https://developer.goto.com/guides/Get%20Started/00_Ref-Get-Started/

This can be done from a Postman request:
1. Go to the "Authorization" tab and select "OAuth 2.0".
2. Check the box that says "Authorize using browser" and use the "Callback URL" as the "Redirect URL" in your GoTo OAuth client configuration.
3. Fill out the section under "Configure New Token" and click "Get New Access Token".

Here is a link to the API docs regarding phone numbers:
https://developer.goto.com/GoToConnect/#operation/getPhoneNumbers
#>

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:failColor = "Red"
    # warning color is yellow, but that is built into Write-Warning
}

function Initialize-AccountInfo
{
    $script:accountKey = "3616656975288904710"
}

function Show-Introduction
{
    Write-Host "This script exports all phone numbers from GoTo Connect." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Prompt-AuthToken
{
    return Read-Host "Please enter your API authorization token"
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function Show-HelpMessage
{
    Write-Host ("To obtain an access token, follow the steps in the getting started and authentication guides: `n" +
        "https://developer.goto.com/guides/Get%20Started/00_Ref-Get-Started/ `n`n" +

        "This can be done from a Postman request: `n" +
        "1. Go to the `"Authorization`" tab and select `"OAuth 2.0`". `n" +
        "2. Check the box that says `"Authorize using browser`" and use the `"Callback URL`" as the `"Redirect URL`" in your GoTo OAuth client configuration. `n" +
        "3. Fill out the section under `"Configure New Token`" and click `"Get New Access Token`". `n") -ForegroundColor $infoColor
}

function Get-PhoneNumbers($accessToken)
{
    $url = "https://api.goto.com/voice-admin/v1/phone-numbers"

    $headers = @{
        Authorization = "Bearer $accessToken"
        Accept = "application/json"
    }

    $queryParams = @{
        accountKey = $accountKey
        pageSize = 100
    }

    $responses = New-Object -TypeName System.Collections.Generic.List[PSObject]
    do
    {
        $response = SafelyInvoke-RestMethod -Uri $url -Method "Get" -Headers $headers -Body $queryParams        
        $responses.Add($response)

        if ($response.nextPageMarker)
        {
            $queryParams["pageMarker"] = $response.nextPageMarker
        }
    }
    while ($response.nextPageMarker)
    
    return $responses
}

function SafelyInvoke-RestMethod($uri, $method, $headers, $body)
{
    try
    {
        $response = Invoke-RestMethod -Uri $uri -Method $method -Headers $headers -Body $body -ErrorVariable "responseError"
    }
    catch
    {
        Write-Host $responseError -ForegroundColor $failColor
        exit
    }

    return $response
}

function Export-PhoneNumbers($accessToken, $apiResponses, [switch]$includeRoutingInfo)
{
    $path = New-DesktopPath -fileName "GoTo phone numbers" -fileExt "csv" -includeTimeStamp
    
    $phoneNumbersProcessed = 0
    foreach ($response in $apiResponses)
    {               
        foreach ($record in $response.items)
        {
            if ($includeRoutingInfo)
            {
                $phoneNumber = Get-PhoneNumberInfo -record $record -includeRoutingInfo
            }
            else
            {
                $phoneNumber = Get-PhoneNumberInfo -record $record
            }
            
            Export-Csv -InputObject $phoneNumber -Path $path -Append -Force -NoTypeInformation

            $phoneNumbersProcessed++
            Write-Progress -Activity "Exporting phone numbers..." -Status "$phoneNumbersProcessed phone numbers processed"
        }
    }
    Write-Host "Finished exporting to $path" -ForegroundColor $successColor
}

function New-DesktopPath($fileName, $fileExt, [switch]$includeTimeStamp)
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")

    if ($includeTimeStamp)
    {
            $timeStamp = (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
        return "$desktopPath\$fileName $timeStamp.$fileExt"
    }
    return "$desktopPath\$fileName.$fileExt"
}

function Get-PhoneNumberInfo($record, [switch]$includeRoutingInfo)
{
    $info = [PSCustomObject]@{
        Name = $record.name
        Number = $record.number
        CallerID = $record.callerIdName
    }
    
    if ($includeRoutingInfo -and $record.routeTo)
    {
        if ($record.routeTo.type -eq "EXTENSION")
        {
            $extension = Get-ExtensionById -accessToken $accessToken -id $record.routeTo.id

            if ($extension.number)
            {
                $routeTo = "$($extension.number) - $($extension.name)"
            }
            else
            {
                $routeTo = $extension.name
            }            
        }
        elseif ($record.routeTo.type -eq "PHONE_NUMBER")
        {
            $routeToPhoneNumber = Get-PhoneNumberById -accessToken $accessToken -id $record.routeTo.id
            $routeTo = "$($routeToPhoneNumber.number) - $($routeToPhoneNumber.name)"
        }

        Add-Member -InputObject $info -Name "RouteTo" -Value $routeTo -MemberType NoteProperty
    }

    return $info
}

function Get-ExtensionById($accessToken, $id)
{
    $url = "https://api.goto.com/voice-admin/v1/extensions/$id"

    $headers = @{
        Authorization = "Bearer $accessToken"
        Accept = "application/json"
    }

    return SafelyInvoke-RestMethod -Uri $url -Method "Get" -Headers $headers
}

function Get-PhoneNumberById($accessToken, $id)
{
    $url = "https://api.goto.com/voice-admin/v1/phone-numbers/$id"

    $headers = @{
        Authorization = "Bearer $accessToken"
        Accept = "application/json"
    }

    return SafelyInvoke-RestMethod -Uri $url -Method "Get" -Headers $headers
}

# main
Initialize-ColorScheme
Initialize-AccountInfo
Show-Introduction
$needHelp = Prompt-YesOrNo "Need help obtaining an access token for the GoTo API?"
if ($needHelp) { Show-HelpMessage }
$accessToken = Prompt-AuthToken
$includeRoutingInfo = Prompt-YesOrNo "Would you like to include routing info? (takes longer)"
$apiResponses = Get-PhoneNumbers $accessToken

if ($includeRoutingInfo)
{
    Export-PhoneNumbers -accessToken $accessToken -apiResponses $apiResponses -includeRoutingInfo
}
else
{
    Export-PhoneNumbers -accessToken $accessToken -apiResponses $apiResponses
}

Read-Host "Press Enter to exit"