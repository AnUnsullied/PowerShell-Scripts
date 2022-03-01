$filename = "FreshworksMobileDeviceAssetExport.csv"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$file = "$path\$filename"

$csvExport = Import-CSV $file


$allMobileDevices = @()


$count = 0
foreach ($item in $csvExport){
    $count +=1
    Write-Output $count

    $mobileDevice = [PSCustomObject]@{
        DisplayName = $item.'Display Name'
        AssetTag = $item.'Asset Tag'
        AssetType = $item.'Asset Type'
        Product = $item.Product
        Division = $item.Division
        AssignedTo = $item.'Assigned To'
        AssetState = $item.'Asset State'
        SerialNumber = $item.'Serial Number'
        CellularLine = $item.'Cellular Line'
        IMEI = $item.IMEI
        ICCID = $item.ICCID
        Carrier = $item.Carrier
        Location = $item.Location
        UsedBy = $item.'Used By'
    }

    if( (-not $mobileDevice.UsedBy) -and ( ($mobileDevice.AssetState -eq 'In Use') -or ($mobileDevice.AssetState -eq 'Reallocated') -or ($mobileDevice.AssetState -eq 'Deployed') ) ){

        $fullName = $mobileDevice.AssignedTo.Split()
        $firstName = $fullName[0]
        $lastName = $fullName[1]

        if(($firstName) -and ($lastName))
        {

            $adUser = Get-ADUser -Filter 'GivenName -eq $firstName -and sn -eq $lastName' -properties EmailAddress | select UserPrincipalName, EmailAddress

            if($adUser.EmailAddress){
                $mobileDevice.UsedBy = $adUser.EmailAddress
                $allMobileDevices += $mobileDevice
            }
        }
    }

}


$allMobileDevices | Export-Csv -Path $path\"mobileDevices.csv" -NoTypeInformation
