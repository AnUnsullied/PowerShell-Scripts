$APIKey = "<API Key goes here!>"
$EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
$HTTPHeaders = @{}
$HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
$HTTPHeaders.Add('Content-Type', 'application/json')


$URL = 'https://company.freshservice.com/api/v2/applications/?per_page=100'
$jsonResponse = (Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders).applications

$allSoftwareAssetID = ""
foreach($softwareAsset in $jsonResponse){
    if($softwareAsset.status -ne 'Managed'){
        $allSoftwareAssetID += $softwareAsset.id
        $allSoftwareAssetID += ","
    }
}


$URL = 'https://company.freshservice.com/api/v2/applications/?ids=' + $allSoftwareAssetID
$jsonResponse = Invoke-RestMethod -Method Delete -Uri $URL -Headers $HTTPHeaders
