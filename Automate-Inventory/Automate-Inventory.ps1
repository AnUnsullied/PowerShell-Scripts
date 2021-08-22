# Determines path of the script
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent

# We create arrays for each Asset Type
# Asset Types are: Accessories, Licenses, Mobile, Monitors, Networking, Printers, Telecom, Workstations
$AssetTypeAccessories = @()
$AssetTypeLicenses = @()
$AssetTypeMobileDevices = @()
$AssetTypeMonitor = @()
$AssetTypeNetworking = @()
$AssetTypePrinter = @()
$AssetTypeTelecom = @()
$AssetTypeComputer = @()


# Convert all .xls files to .csv in the working folder
foreach($file in (Get-ChildItem -Path "$path\Working\")){
    
    # Get the date in a specific format so we can use it for the files
    $workingDate = Get-Date -format yyyyMMdd
    # Create a new file name
    $newFileName = "$path\Completed\" + $workingDate +".csv"

    # Create Excel Workbook to open the .xls file and save as a .csv file in the Completed folder
    $ExcelWB = New-Object -ComObject excel.application
    $Workbook = $ExcelWB.Workbooks.Open($file.FullName)
    $Workbook.SaveAs($newFileName, 6)
    $Workbook.Close($False)
    $ExcelWB.quit()

    # We don't need the old .xls file since it's been converted to .csv
    Remove-Item $file.FullName
}

# Gets all the csv files in the folder without blanks
$ImportFiles = Import-Csv -Path (Get-ChildItem -Path $newFileName -Filter '*.csv').FullName | Where 'Asset Type' -ne ""




# Iterate through each entry in the .csv file
foreach($item in $ImportFiles){
    
    # Switch statement to determine what the Asset Type matches so we can add it to the array
    Switch($item.'Asset Type'){
        "Accessories" {
            $AssetTypeAccessories += $item
        }
        "Licenses" {
            $AssetTypeLicenses += $item
        }
        "Mobile" {
            $AssetTypeMobileDevices += $item
        }
        "Monitor" {
            $AssetTypeMonitor += $item
        }
        "Networking" {
            $AssetTypeNetworking += $item
        }
        "Printer" {
            $AssetTypePrinter += $item
        }
        "Telecom" {
            $AssetTypeTelecom += $item
        }
        "Computer" {
            $AssetTypeComputer += $item
        }
        default {
            Write-Output "______"
            Write-Output "ERROR: Incorrect Asset Type: $item.'Asset Type'"
            Write-Output $item
            Write-Output "______"
        }
    }
    
}

# Create file names for each Asset Type
$fileNameAccessories = (Get-Date -format yyyyMMdd) + "AssetTypeAccessories.csv"
$fileNameLicenses = (Get-Date -format yyyyMMdd) + "AssetTypeLicenses.csv"
$fileNameMobileDevices = (Get-Date -format yyyyMMdd) + "AssetTypeMobileDevices.csv"
$fileNameMonitor = (Get-Date -format yyyyMMdd) + "AssetTypeMonitor.csv"
$fileNameNetworking = (Get-Date -format yyyyMMdd) + "AssetTypeNetworking.csv"
$fileNamePrinter = (Get-Date -format yyyyMMdd) + "AssetTypePrinter.csv"
$fileNameTelecom = (Get-Date -format yyyyMMdd) + "AssetTypeTelecom.csv"
$fileNameComputer = (Get-Date -format yyyyMMdd) + "AssetTypeComputer.csv"

# Call each Asset Type array and pipe an export to a .csv file
$AssetTypeAccessories | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameAccessories"
$AssetTypeLicenses | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameLicenses"
$AssetTypeMobileDevices | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameMobileDevices"
$AssetTypeMonitor | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameMonitor"
$AssetTypeNetworking | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameNetworking"
$AssetTypePrinter | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNamePrinter"
$AssetTypeTelecom | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameTelecom"
$AssetTypeComputer | Export-CSV -NoTypeInformation -Path "$path\Completed\$fileNameComputer"

