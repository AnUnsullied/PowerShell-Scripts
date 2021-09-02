<#
    This script looks for a file called "InventoryFile.xlsx" in the folder where the script itself is stored.
    It creates a .csv file of all the worksheet in the InventoryFile.xlsx workbook
    It then looks for deployed entries with other criteas and process those entries:
        Sorts by location
        Loops through each entry
            Determines all entries belonging to a property
            Determines the cost of the entry based off of the Order Tracker sheet
            Copies the Pay Request Form file and fills that out with the entries
#>




# Converts any worksheet in an Excel workbook to a .csv file
function ExportWSToCSV ($excelFileName, $csvLoc, $path){
    $excelFile = "$path\" + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $n = (Get-Date -format yyyyMMdd) + "_" + $excelFileName + "_" + $ws.Name
        $ws.SaveAs($csvLoc + $n + ".csv", 6)
    }
    $E.Quit()
}




# Calculates and finds the cost of a PO if it is in a csv file
function Get-PO-Cost($PO, $item, $path){
    $cost = 0.00

    $ImportFilePO = Import-Csv -Path ("$path\" + (Get-Date -format yyyyMMdd) + "_" + "InventoryFile_Order Tracker.csv") | Sort-Object 'PO / Order #'

    foreach ($PurchaseOrder in $ImportFilePO){
        
        if( ($PO -eq $PurchaseOrder.'PO / Order #') -and ($item -eq $PurchaseOrder.'Description') ){
            $cost = (([decimal]($PurchaseOrder.'Unit Cost'.Replace('$', '').Replace(',', ''))) + ([decimal]($PurchaseOrder.'Unit Tax (Auto)'.Replace('$', '').Replace(',', ''))) )
        }

    }

    return $cost
}





# shortens the name of the property
function Clean-Property-Name($propertyName){
    $propertyName = $propertyName.Replace('Apartments and Townhomes', '')
    $propertyName = $propertyName.Replace('Apartments', '')
    $propertyName = $propertyName.Replace('Townhomes', '')
    $propertyName = $propertyName.Trim()
    $propertyName = $propertyName.Replace(' ', '_')

    return $propertyName
}


# Copies the Pay Request file and renames the new file accordingly with date and property
function Create-Copied-Pay-Request-File($propertyName, $payRequestYear, $payRequestMonth, $path){

    # Filename cannot be longer than 31 characters
    $propertyName = Clean-Property-Name -propertyName $propertyName

    $newFilename = $payRequestYear + "_" + $payRequestMonth + "-" + $propertyName + ".xlsx"

    Copy-Item "$path\Do Not Use - IT_PayRequestForm.xlsx" -Destination $path\$newFilename

    return "$path\$newFilename", $propertyName
}


# returns the corresponding string month based on digit version of the month
function Get-Month($payRequestMonth){

    switch($payRequestMonth){
        ("1"){
            return "January"
        }
        ("2"){
            return "February"
        }
        ("3"){
            return "March"
        }
        ("4"){
            return "April"
        }
        ("5"){
            return "May"
        }
        ("6"){
            return "June"
        }
        ("7"){
            return "July"
        }
        ("8"){
            return "August"
        }
        ("9"){
            return "September"
        }
        "10"{
            return "October"
        }
        "11"{
            return "November"
        }
        "12"{
            return "December"
        }
        default{
            return $payRequestMonth
        }
    }


}


# Creates the Pay Request and opens up Excel to edit the files
function Create-Pay-Request($propertyName, $deployedItems, $path, $payRequestYear, $payRequestMonth, $payRequestID){

    # Calls function to create the pay request file; receive the file path
    $payRequestFile, $propertyname = Create-Copied-Pay-Request-File -propertyName $propertyName -payRequestYear $payRequestYear -payRequestMonth $payRequestMonth -path $path

    # Gets the string version of the month based off of the digit
    $payRequestMonth = Get-Month -payRequestMonth $payRequestMonth

    # Cleans up the name of the property (shortens it)
    $propertyName = Clean-Property-Name -propertyName $propertyName

    

    # Opens Excel object, finds workbook and worksheet
    $excel = New-Object -ComObject Excel.Application
    #$excel.Visible = $true
    $workbook = $excel.Workbooks.Open($payRequestFile)
    $sheet = $workbook.Worksheets.Item(1)



    # Update the Project with Property Name (D,3)
    $sheet.Cells.Item(3,4) = $propertyName
    
    # Update the Type of Work field (G, 4)
    $sheet.Cells.Item(3,7) = ("$($PayRequestYear) of $($payRequestMonth) ")

    # Update the Pay Request ID (I, 2)
    $sheet.Cells.Item(2,9) = $payRequestID

    # Index where the Y axis starts before we insert the asset to request payment for
    $startIndexY = 7

    # Loop through the list
    foreach($asset in $deployedItems){
    

        # Update quantity field
        $sheet.Cells.Item($startIndexY,2) = $asset.assetQuantity

        # Update Unit Cost field
        $sheet.Cells.Item($startIndexY,3) = $asset.assetPOCost
        
        # Update Category field
        $sheet.Cells.Item($startIndexY,5) = $asset.assetType

        # Update Description field
        $sheet.Cells.Item($startIndexY,6) = $asset.assetProduct

        # Update UID field
        if($asset.assetTag){
            $sheet.Cells.Item($startIndexY,7) = $asset.assetTag
        }
        else{
            $sheet.Cells.Item($startIndexY,7) = "N/A"
        }

        # Update Notes field
        $sheet.Cells.Item($startIndexY,9) = $asset.assetReference

        #Increment
        $startIndexY += 1

    }

    $workbook.Save()
    $excel.Quit()

    return $payRequestID
}




# function to grab user's input and return the values
function Get-User-Input(){
    $payRequestYear = Read-Host "Input the Pay Request Year you want (XXXX)"
    $payRequestMonth = Read-Host "Input the Pay Request Month you want (1-12)"
    $payRequestID = Read-Host "Input the previously used Pay Request ID"
    return $payRequestYear, $payRequestMonth, $payRequestID
}




# Variables to be used
$placeholderProperty = "" # Placeholder property attribute
$placeholderItems = @() # Placeholder items array


# Determines path of the script
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent



# Call function to get input about year and month we are looking at
$payRequestYear, $payRequestMonth, $payRequestID = Get-User-Input


# Call function to create CSV file of all worksheets in Inventory 2.0XX
ExportWSToCSV -excelFileName "InventoryFile" -csvLoc "$path\" -path $path


# Gets all the entries in the DeploymentTrack sheet (now a csv file) where it is a deployed object, and sorts it by location
$ImportFileDeployment = Import-Csv -Path ("$path\" + (Get-Date -format yyyyMMdd) + "_" + "InventoryFile_DeploymentTracking.csv") | Where-Object {$_.'Asset State' -eq "Deployed"} | Sort-Object Location




# Loop through that .csv file
foreach ($item in $ImportFileDeployment){

    # Split string date into values we can use later (year, day, month)
    $itemDate = $item.Date.Split('/')

    # Set Variables for readibility
    $asset = [PSCustomObject]@{
        'assetState' = $item.'Asset State'
        'assetTag' = $item.'Asset Tag'
        'assetType' = $item.'Asset Type'
        'assetDate' = $item.Date
        'assetYear' = $itemDate[2]
        'assetMonth' = $itemDate[0]
        'assetDay' = $itemDate[1]
        'assetDeviceRole' = $item.'Device Role'
        'assetDisplayName' = $item.'Display Name'
        'assetICCID' = $item.ICCID
        'assetIMEI' = $item.IMEI
        'assetInitials' = $item.Initials
        'assetLocation' = $item.Location
        'assetPhoneNumber' = $item.'Phone Number'
        'assetPO' = $item.PO
        'assetProduct' = $item.Product
        'assetQuantity' = $item.Quantity
        'assetReference' = $item.Reference
        'assetWWD' = $item.WWD

        'assetPOCost' = Get-PO-Cost -PO $item.PO -item $item.Product -path $path
    } 

    # Conditions for pay requests -> year, month must equal user input. Pay Request must not contain the words billed to site
    if( ($asset.assetYear -eq $payRequestYear) -and ($asset.assetMonth -eq $payRequestMonth) -and ( -not ($asset.assetPO.Contains("Billed to Site"))) ){
        
        # If the placeholder property exists in the entry (i.e. this isn't the very first run)
        If ($placeholderProperty){

            # We add to the existing list of devices at the placeholder property
            If($placeholderProperty.Equals($asset.assetLocation)){
                
                # Expand and add the current device to the list
                $placeholderItems += $asset
            }

            # Does not match previous property, means this entry is from a new propery. We finish up the previous property pay request then process the current one
            Else{

                # We increment the previous Pay Request ID (since it has to be unique)
                $payRequestID = [string]([int]($payRequestID) + 1)
                
                # Calls function to create Pay Requests
                Create-Pay-Request -propertyName $placeholderProperty -deployedItems $placeholderItems -path $path -payRequestYear $payRequestYear -payRequestMonth $payRequestMonth -payRequestID $payRequestID



                # Previous property Pay Requests are now all set, move on to initalize pay request for current property

                # Now we update the current placeholder property to the new attributes of this device
                $placeholderProperty = $asset.assetLocation
                # Clear the array and variables
                $placeholderItems = @()
                # Add the device to the list of the placeholder devices
                $placeholderItems += $asset
            }

        }

        # PlaceholderProperty doesn't exist; this is the first run
        else {
        
            # intialize values
            $placeholderProperty = $asset.assetLocation
            $placeholderItems += $asset

        }



    }
    
}


# Process the last property if exists
if ($placeholderItems){
    Create-Pay-Request -propertyName $placeholderProperty -deployedItems $placeholderItems -path $path -payRequestYear $payRequestYear -payRequestMonth $payRequestMonth -payRequestID $payRequestID
}

# Clear the array and variables
$placeholderProperty = ""
$placeholderItems = @()
