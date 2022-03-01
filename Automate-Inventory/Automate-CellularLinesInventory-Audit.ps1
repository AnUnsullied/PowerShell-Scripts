<#
The script finds the following:
1. Discrepencies between IMEI, ICCID, Location, and Division
2. Any connections with a line to an EOL / Retired status
3. Lines with no relationships that are active
4. Lines that are suspended but has a relationship

#>

# Function: Gets asset based on its asset tag
# Usage Example:
# Get-Asset-by-AssetTag "5178989957"
function Get-Asset-by-AssetTag($assetTag){
    # Build the querry we want (asset tag)
    $querry = "`"asset_tag%3A%27" + $assetTag + "%27`""

    # Create the URL for the API Call
    $URL = 'https://company.freshservice.com/api/v2/assets?search=' + $querry

    # return the results of the API call
    $outputJSON = Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders

    
    return $outputJSON.assets
}

# Function: Gets asset based on display ID
function Get-Asset-by-DisplayID($assetDisplayID){
    # Build the URL
    $URL = 'https://company.freshservice.com/api/v2/assets/' + $assetDisplayID + '?include=type_fields' # View specific item

    return Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders
}

#Function: Get asset relationship based on asset display ID
function Get-Asset-Relationship($assetDisplayID){
    # Build the URL
    $URL = 'https://company.freshservice.com/api/v2/assets/' + $assetDisplayID + '/relationships'
    $outputJSON = Invoke-RestMethod -Method GET -Uri $URL -Headers $HTTPHeaders

    return $outputJSON
}

#Function: Get asset device asset based on phone number it has a relationship with
function Get-Associated-Asset-by-CellularLine($cellularLine){
    # Cell line should be already used as an asset tag
    $asset = Get-Asset-by-AssetTag -assetTag $cellularLine

    $relationship = Get-Asset-Relationship $asset.display_id

    return (Get-Asset-by-DisplayID $relationship.relationships.primary_id).asset
}


function Get-Location-by-ID($locationID){
    $URL = 'https://company.freshservice.com/api/v2/locations/' + $locationID

    return Invoke-RestMethod -Method GET -Uri $URL -Headers $HTTPHeaders
}


function Get-AssetType-by-ID($assetTypeID){
    $URL = 'https://company.freshservice.com/api/v2/asset_types/' + $assetTypeID

    return Invoke-RestMethod -Method GET -Uri $URL -Headers $HTTPHeaders
}


$APIKey = "<API Key goes here>"
$EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
$HTTPHeaders = @{}
$HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
$HTTPHeaders.Add('Content-Type', 'application/json')



$filename = "FreshworksExportofCellularLines.csv"


$counterThreshold = 300
$counter = 0
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent


$file = "$path\$filename"
$csvExport = Import-CSV $file


$issuesDivision_Location = @()
$issuesIMEI_ICCID = @()
$issuesEOL_RetiredAssets = @()
$issuesNoRelationships = @()
$issuesMissingRelationships = @()

foreach($entry in $csvExport){

    Write-Output $entry.'Cellular Line'

    # Find the asset the cellular line has a relationship to
    $assetDevice = Get-Associated-Asset-by-CellularLine -cellularLine $entry.'Cellular Line'
    $counter += 3 # Increment the API call we did
    
    
    if($counter -ge $counterThreshold){
        Write-Output "API Call threshold has been reached. Pausing for 60 seconds."
        Start-Sleep -Seconds 30
        
        Write-Output "30 seconds left."
        Start-Sleep -Seconds 20

        Write-Output "10 seconds left."
        Start-Sleep -Seconds 10

        $counter = 0
    }

    # If the line is in use (it should have a relationship)
    if($entry.Status -eq 'Active'){

        if(-not ($assetDevice)){
            # if the asset does not exist despite being an active line
            $issuesMissingRelationships += $entry
        }




        $assetState = $assetDevice.type_fields.asset_state_15000469173

        # Does the asset the line has a relationship with is in a EOL / Retired status?
        if( ($assetState -eq "EOL - Out of Support") -or ($assetState -eq "EOL - Damaged") -or ($assetState -eq "Retired - Sold") -or ($assetState -eq "Retired - Lost") -or ($assetState -eq "Retired - Disposed") ){
            # If so, this is incorrect and needs to be audited
            $issuesEOL_RetiredAssets += $entry
        }


        # Does the division for the cellular line matches the division for the related asset?
        $divisionFlag = ($entry.Division -eq $assetDevice.type_fields.division_15000469173)
        # Does the Location for the cellular line matches the Location for the related asset?
        $locationFlag = ($entry.Location -eq (Get-Location-by-ID -locationID $assetDevice.location_id).location.name)
        $counter += 1   # Increment the API call used
        
        # If the division and / or location flag does not match, add the entry to the array (will be turned to .csv file later)
        if(-not ($divisionFlag -and $locationFlag)){
            $issuesDivision_Location += $entry
        }

       


        # Need to compare ICCID & IMEI, but it's unique ID in FW to the assets themselves.
        if($entry.'Intended Device Type' -eq "Network"){
            # Get ICCID & IMEI for network devices

            # What network devices is this? Firewall, Cellular Gateway, USB Hotspots?
            $assetDeviceType = (Get-AssetType-by-ID -assetTypeID $assetDevice.asset_type_id).asset_type.name
            $counter += 1

            switch($assetDeviceType){
                "Firewall"{
                    $assetDeviceICCID = $assetDevice.type_fields.iccid_15000469189
                    $assetDeviceIMEI = $assetDevice.type_fields.imei_15000469189

                    Break
                }
                "Cellular Gateway"{
                    $assetDeviceICCID = $assetDevice.type_fields.iccid_15000940241
                    $assetDeviceIMEI = $assetDevice.type_fields.imei_15000940241
                    
                    Break
                }
                "USB Hotspot"{
                    $assetDeviceICCID = $assetDevice.type_fields.iccid_15000940245
                    $assetDeviceIMEI = $assetDevice.type_fields.imei_15000940245
                    
                    Break
                }
                default{
                    $assetDeviceICCID = ""
                    $assetDeviceIMEI = ""
                }
            }

        }
        elseif($entry.'Intended Device Type' -eq "Computer"){
            # Get ICCID & IMEI for Computer

            $assetDeviceICCID = $assetDevice.type_fields.iccid_15000469178
            $assetDeviceIMEI = $assetDevice.type_fields.imei_15000469178

        }
        else{
            # Get ICCID & IMEI for Mobile Devices

            $assetDeviceICCID = $assetDevice.type_fields.iccid_15000469181
            $assetDeviceIMEI = $assetDevice.type_fields.imei_15000469181
        }



        # Does the ICCID for the cellular line matches the ICCID for the related asset?
        $ICCIDFlag = ($entry.ICCID -eq $assetDeviceICCID)
        # Does the IMEI for the cellular line matches the IMEI for the related asset?
        $IMEIFlag = ($entry.IMEI -eq $assetDeviceIMEI)


        # If the IMEI and / or ICCID flag does not match, add the entry to the array (will be turned to .csv file later) 
        if(-not ($ICCIDFlag -and $IMEIFlag)){
            $issuesIMEI_ICCID += $entry
        }



    }
    else{
        # Line is suspended, canceled (should NOT have a relationship)
        if($assetDevice){
            # if the asset does exist (meaning it had a relationship despite being a suspended or canceled line)
            $issuesNoRelationships += $entry
        }
    }


    
}

$issuesDivision_Location | Export-Csv -Path $path\"DiscrepenciesDivision_Location.csv" -NoTypeInformation
$issuesIMEI_ICCID | Export-Csv -Path $path\"DiscrepenciesIMEI_ICCID.csv" -NoTypeInformation
$issuesEOL_RetiredAssets | Export-Csv -Path $path\"DiscrepenciesEOL_RetiredAssets.csv" -NoTypeInformation
$issuesNoRelationships | Export-Csv -Path $path\"DiscrepenciesNoRelationships.csv" -NoTypeInformation
$issuesMissingRelationships | Export-Csv -Path $path\"DiscrepenciesMissingRelationships.csv" -NoTypeInformation
