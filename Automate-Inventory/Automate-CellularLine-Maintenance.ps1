# Function: Gets asset based on its asset tag
function Get-Asset-by-AssetTag($assetTag){
    # Build the querry we want (asset tag)
    $querry = "`"asset_tag%3A%27" + $assetTag + "%27`""

    # Create the URL for the API Call
    $URL = 'https://company.freshservice.com/api/v2/assets?search=' + $querry

    return Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders
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

    return (Invoke-RestMethod -Method GET -Uri $URL -Headers $HTTPHeaders)
}

function Get-Location-by-ID($locationID){
    $URL = 'https://company.freshservice.com/api/v2/locations/' + $locationID

    return Invoke-RestMethod -Method GET -Uri $URL -Headers $HTTPHeaders
}

function Delete-Relationship($relationshipID){
    $URL = 'https://company.freshservice.com/api/v2/relationships?ids=' + $relationshipID
    $outputJSON = Invoke-RestMethod -Method DELETE -Uri $URL -Headers $HTTPHeaders
}

# Function: Creates a Uses relationship between a mobile device asset and a cellular line asset
# Usage Example:
# Create-Asset-Uses-Relationships -primaryID $mobileDeviceDisplayID -secondaryID $cellularNumberDisplayID
function Create-Asset-Uses-Relationships($primaryID, $secondaryID){
    # URL to create relationships for REST API calls
    $URL = 'https://company.freshservice.com/api/v2/relationships/bulk-create'

    $relationshipTypeIDUses = 15000239515 # Relationship Type 'Uses' ID
    # We'll only be using asset type, not 
    $primaryType = "asset"
    $secondaryType = "asset"

    # Relationship object
    $relationshipAttribute = @{}
    $relationshipAttribute.Add('relationship_type_id', $relationshipTypeIDUses)
    $relationshipAttribute.Add('primary_id', $primaryID)
    $relationshipAttribute.Add('primary_type', $primaryType)
    $relationshipAttribute.Add('secondary_id', $secondaryId)
    $relationshipAttribute.Add('secondary_type', $secondaryType)

    # Array for all relationship changes for the bulk create
    $relationshipAll = @($relationshipAttribute)

    $relationshipAttribute = @{'relationships' = $relationshipAll}
    $inputJSON = $relationshipAttribute | ConvertTo-Json

    $outputJSON = Invoke-RestMethod -Method Post -Uri $URL -Headers $HTTPHeaders -Body $inputJSON
}

function Update-Cellular-Line-Asset($assetDisplayID, $division, $location_id, $ICCID, $IMEI, $status){

    # URL needed to pull existing asset in Freshworks so we can get the attributes
    $URL = 'https://company.freshservice.com/api/v2/assets/' + $assetDisplayID + '?include=type_fields'
    # Result of the API call
    $outputJSON = Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders

    # Set attributes we are not changing to a variable so we can create the updated asset
    $name = $outputJSON.asset.name
    $asset_type_id = $outputJSON.asset.asset_type_id
    $asset_tag = $outputJSON.asset.asset_tag
    $phoneNumber = $outputJSON.asset.type_fields.phone_number_15000936857
    $carrier = $outputJSON.asset.type_fields.carrier_15000936857
    $intendedDeviceType = $outputJSON.asset.type_fields.intended_device_type_15000936857
    
    if( (-not ($division)) ) {$division = $outputJSON.asset.type_fields.division_15000931793}
    if( (-not ($location_id)) ) {$location_id = $outputJSON.asset.location_id}
    if( (-not ($ICCID)) ) {$ICCID = $outputJSON.asset.type_fields.iccid_15000936857}
    if( (-not ($IMEI)) ) {$IMEI = $outputJSON.asset.type_fields.imei_15000936857}
    if( (-not ($status)) ) {$status = $outputJSON.asset.type_fields.status_15000936857}

    
    # URL needed to update asset
    $URL = 'https://company.freshservice.com/api/v2/assets/' + $assetDisplayID

    $ConfigItems = @{}
    $ConfigItems.Add('name', $name)
    $ConfigItems.Add('asset_type_id', $asset_type_id)
    $ConfigItems.Add('asset_tag', $asset_tag)
    $ConfigItems.Add('location_id', $location_id)

    $customItems = @{}
    $customItems.Add('division_15000931793', $division)
    $customItems.Add('phone_number_15000936857', $phoneNumber)
    $customItems.Add('iccid_15000936857', $ICCID)
    $customItems.Add('imei_15000936857', $IMEI)
    $customItems.Add('status_15000936857', $status)
    $customItems.Add('carrier_15000936857', $carrier)
    $customItems.Add('intended_device_type_15000936857', $intendedDeviceType)

    $ConfigItems.Add('type_fields', $customItems)


    $ConfigItems = @{'asset' = $ConfigItems}
    $inputJSON = $ConfigItems | ConvertTo-Json


    $result = Invoke-RestMethod -Method Put -Uri $URL -Headers $HTTPHeaders -Body $inputJSON
} 

# Function: Gets asset based on its asset tag
function Get-Location-by-Name($locationName){
    # Build the querry we want (asset tag)

    $locationName = $locationName -replace(" ", "%20")

    $query = "`"name%3A%27" + $locationName + "%27`""

    # Create the URL for the API Call
    $URL = 'https://company.freshservice.com/api/v2/locations?query=' + $query

    return Invoke-RestMethod -Method Get -Uri $URL -Headers $HTTPHeaders
}




function Suspend-Cellular-Line-Asset($cellularAsset){
    
    $locationID = (Get-Location-by-Name -locationName "Company Location").locations.id

    Update-Cellular-Line-Asset -assetDisplayID $cellularAsset.display_id -division "Company Division" -location_id $locationID -status "Suspended"

    $body = "The following cellular line should be suspended.`n"
    $body += $cellularAsset


    Send-MailMessage -From "companyIT@ecompany.com" -To "me@ecompany.com" -SmtpServer "serever.company.com" -Subject "Suspend Cellular Line" -Body $body
}


function Cancel-Cellular-Line-Asset($cellularAsset){
    

    Update-Cellular-Line-Asset -assetDisplayID $cellularAsset.display_id -status "Cancelled"

    $body = "The following cellular line should be cancelled.`n"
    $body += $cellularAsset


    Send-MailMessage -From "companyIT@ecompany.com" -To "me@ecompany.com" -SmtpServer "serever.company.com" -Subject "Cancel Cellular Line" -Body $body
}








$APIKey = "<API Key goes here!>"
$EncodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $APIKey,$null)))
$HTTPHeaders = @{}
$HTTPHeaders.Add('Authorization', ("Basic {0}" -f $EncodedCredentials))
$HTTPHeaders.Add('Content-Type', 'application/json')






$cellularLine = Read-Host "Enter a cellular line"

$cellularAsset = Get-Asset-by-AssetTag -assetTag $cellularLine


if($cellularAsset.assets.Count -eq 1){
    # Asset was found in FW, and is the only asset found

    $cellularAsset = (Get-Asset-by-DisplayID -assetDisplayID $cellularAsset.assets.display_id).asset

    $cellularRelationship = Get-Asset-Relationship -assetDisplayID $cellularAsset.display_id


    if($cellularRelationship.relationships){
        $associatedAsset = (Get-Asset-by-DisplayID $cellularRelationship.relationships.primary_id).asset

        $location = (Get-Location-by-ID -locationID $associatedAsset.location_id).location.name
        if($cellularAsset.type_fields.contract_end_date_15000936857){
            $endDate = $cellularAsset.type_fields.contract_end_date_15000936857
        }
        else{
            $endDate = "N/A"
        }
        $outStr = "Cellular line " + $cellularLine + " with contract end date " + $endDate + " is being used by: " + $associatedAsset.name + " at: " + $location + "`n"
        Write-Output $outStr
        

        Write-Host "What do you want to do with this cellular line?"
        Write-Host "1. Delete the existing relationship and create a new relationship?"
        Write-Host "2. Delete the existing relationship and suspend this line?"
        Write-Host "3. Delete the existing relationship and cancel this line?"
        Write-Host "4. Do nothing.`n"

        $userChoice = Read-Host "Enter your choice"
        switch($userChoice){
            "1"{
                $deleteChoice = Read-Host "Are you sure you want to delete the existing relationship (y/n)"
                if($deleteChoice -eq "y"){
                    $outStr = "Relationship between " + $cellularLine + " and " + $associatedAsset.name + " has been deleted.`n"
                    Write-Output $outStr

                    Delete-Relationship -relationshipID $cellularRelationship.relationships.id



                    $deviceAssetTag = Read-Host "Enter the asset tag of the device you want $cellularLine to be associated with"
                    $deviceAsset = Get-Asset-by-AssetTag -assetTag $deviceAssetTag
                    if($deviceAsset.assets.Count -eq 1){
                        # Asset was found in FW, and is the only asset found

                        $deviceAsset = (Get-Asset-by-DisplayID -assetDisplayID $deviceAsset.assets.display_id).asset
                        $outStr = "Creating relationship between " + $cellularLine + " and asset tag " + $deviceAssetTag + " (" + $deviceAsset.type_fields.serial_number_15000469173 + ")`n"
                        Write-Output $outStr


                        Create-Asset-Uses-Relationships -primaryID $deviceAsset.display_id -secondaryID $cellularAsset.display_id
                    }
                    else{
                        Write-Output "Multiple assets found or no assets found. Cannot establish relationship.`n"
                    }
                }
                
                break;
            }
            "2"{
                $deleteChoice = Read-Host "Are you sure you want to delete the existing relationship (y/n)"
                if($deleteChoice -eq "y"){
                    $outStr = "Relationship between " + $cellularLine + " and " + $associatedAsset.name + " has been deleted."
                    Write-Output $outStr

                    Delete-Relationship -relationshipID $cellularRelationship.relationships.id

                    Suspend-Cellular-Line-Asset -cellularAsset $cellularAsset


                }
                
                break;
            }
            "3"{
                $deleteChoice = Read-Host "Are you sure you want to delete the existing relationship (y/n)"
                if($deleteChoice -eq "y"){
                    $outStr = "Relationship between " + $cellularLine + " and " + $associatedAsset.name + " has been deleted."
                    Write-Output $outStr

                    Delete-Relationship -relationshipID $cellularRelationship.relationships.id

                    
                    Cancel-Cellular-Line-Asset -cellularAsset $cellularAsset
                }

                break;
            }
        }


    }
    else{
        if($cellularAsset.type_fields.contract_end_date_15000936857){
            $endDate = $cellularAsset.type_fields.contract_end_date_15000936857
        }
        else{
            $endDate = "N/A"
        }
        $outStr = "This cellular line " + $cellularLine + " with contract end date " + $endDate + " has no relationships.`n"
        Write-Output $outStr

        Write-Host "What do you want to do with this cellular line?"
        Write-Host "1. Suspend the line in FW and generate a ticket to suspend?"
        Write-Host "2. Cancel the line in FW and generate a ticket to cancel?"
        Write-Host "3. Associate the line with a new asset."
        Write-Host "4. Do nothing.`n"

        $userChoice = Read-Host "Enter your choice"
        switch($userChoice){
            "1"{
                $outStr = "This cellular line " + $cellularLine + " has been suspended in FW and a ticket has been made to suspend the line with the provider.`n"
                Write-Output $outStr
                Suspend-Cellular-Line-Asset -cellularAsset $cellularAsset

                break;
            }
            "2"{
                $outStr = "This cellular line " + $cellularLine + " has been cancelled in FW and a ticket has been made to cancel the line with the provider.`n"
                Write-Output $outStr
                Cancel-Cellular-Line-Asset -cellularAsset $cellularAsset

                break;
            }
            "3"{
                $deviceAssetTag = Read-Host "Enter the asset tag of the device you want $cellularLine to be associated with"
                $deviceAsset = Get-Asset-by-AssetTag -assetTag $deviceAssetTag
                if($deviceAsset.assets.Count -eq 1){
                    # Asset was found in FW, and is the only asset found

                    $deviceAsset = (Get-Asset-by-DisplayID -assetDisplayID $deviceAsset.assets.display_id).asset
                    $outStr = "Creating relationship between " + $cellularLine + " and asset tag " + $deviceAssetTag + " (" + $deviceAsset.type_fields.serial_number_15000469173 + ")`n"
                    Write-Output $outStr


                    Create-Asset-Uses-Relationships -primaryID $deviceAsset.display_id -secondaryID $cellularAsset.display_id
                }
                else{
                    Write-Output "Multiple assets found or no assets found. Cannot establish relationship.`n"
                }
            }
        }
    }
}
elseif($cellularAsset.assets.Count -gt 1){
    # Too many assets found
    Write-Output "Multiple assets found. Potential duplicates when searching for: $cellularLine"
}
else{
    # Asset was not found in FW
    Write-Output "The cellular line was not found in FW."

    #Create new asset
}
