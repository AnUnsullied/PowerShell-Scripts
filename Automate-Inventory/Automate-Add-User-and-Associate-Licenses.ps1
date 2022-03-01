$APIKey = "<API Key goes here>"
$EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
$HTTPHeaders = @{}
$HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
$HTTPHeaders.Add('Content-Type', 'application/json')

function Get-UserID-by-Email($email){
    # We check requester users first
    $URL = 'https://company.freshservice.com/api/v2/requesters?query="primary_email:%27'
    
    # Need to format the email address for the query
    $email = $email.Replace('@', '%40')
    $URL += $email + "%27`""

    # Get our json response after searching for the requester
    $jsonResponse = (Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders)


    # Does the requester exist / is found? If no, then results are 0
    if($jsonResponse.requesters.Count -eq 0){
        # If no requester users is found, they might be an agent user.


        # We create the URL for the api call to find an agent user
        $URL = 'https://company.freshservice.com/api/v2/agents?query="email:%27' + $email + "%27`""


        $jsonResponse = (Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders)

        # Error checking - if no requester or agent user found, we have a user not in the system
        if($jsonResponse.agents.Count -eq 0){
            return -1
        }

        # Agent user found; returning their ID
        return $jsonResponse.agents.id
    }

    # Return requester user id when found
    return $jsonResponse.requesters.id   
}





$filename = "softwareuserlistexport.csv"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$file = "$path\$filename"

$csvExport = Import-CSV $file

# Contact 19, Software 15030898866



$applicationUsers = @()
$licenseID = 111 # Contract display ID in FW
$softwareID = 111111111111111 # Software display ID in FW


foreach($entry in $csvExport){
    $appUser = @{}
    $appUser.Add('user_id', (Get-UserID-by-Email -email $entry.Email))
    $appUser.Add('license_id', $licenseID)

    $applicationUsers += $appUser
}

$application = @{'application_users' = $applicationUsers}
$inputJSON = $application | ConvertTo-Json

$URL = 'https://company.freshservice.com/api/v2/applications/' + $softwareID + '/users'
$outputJson = Invoke-RestMethod -Method Post -Uri $URL -Headers $HTTPHeaders -Body $inputJSON
