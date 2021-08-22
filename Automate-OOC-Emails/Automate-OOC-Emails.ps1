# This script looks into a .csv file that contains information about devices that are considered Out of Contact (OOC) for a propety. The email is sent to the manager of that property


# This function generates and sends the OOC email
function Send-OOC-Email($placeholderItems, $placeholderProperty){

    # Creates a new object so we can create the emails
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    # Looks into AD to find a user located at the property and if they have a Manager job title
    $manager = Get-ADUser -Filter {Office -eq $placeholderProperty -and Title -eq "Manager" -and Enabled -eq $true} -Properties *

    # If we find a user that matches that criteia, we add them as a recipient to the email
    if($manager){
        $mail.Recipients.Add($manager.Mail) > $null
    }

    # IT team should be tagged in the email ideally
    $mail.Recipients.Add("IT-Team@example.com") > $null


    # CSS style so we can use for the email body
    $Style = '<style> table, th, td{ border: 1px solid; } </style> '

    # Based off of the data from our table, we can call ConvertTo-HTML to format the data in the email
    $table = [PSCustomobject]$placeholderItems | ConvertTo-Html -Fragment -As Table
    $htmlTable = ConvertTo-Html -Head $Style -Body $table


    # HTML Email
    $from    =  "this-is-yourself@example.com"
    $subject =  "Out of Contact Device Audit: [$placeholderProperty]"

    $body = "Hello!"
    $body += "<br />"
    $body += "<br />"
    $body += "We have detected that the following device(s) are out of contact, meaning that they have not checked in our system in at least over <font color=#FF0000> 30 days </font> or more."
    $body += "<br />"
    $body += "It appears you are the best contact for this audit, so please review the information for the following device(s):"
    $body += "<br />"
    $body += "<br />"
    $body += $htmlTable
    $body += "<br />"
    $body += "If you have the device(s) in your possession, is the device(s) needed at the site?"
    $body += "<ul> <li>If yes, we urge you to keep the device powered on, charged, and connected as if it was a laptop so that it remains connected to our mobile device management system. Please ensure the device(s) gets connected immediately.</li>"
    $body += "<li>If not, please let us know so  we can provide you a return label for this tablet</li></ul>"
    $body += "If you do not have the device(s) in your possession, please let us know immediately. We'll consider the device(s) as missing, and if there is an active cellular plan, we will need to cancel it on our end."
    $body += "<br />"
    $body += "<br />"
    $body += "<b><font color=#FF0000>NOTE:</b></font> If we do not get any sort of response from you within <b><i>three</b></i> business days, we will consider the device(s) as missing and will cancel the line. To ensure that the cellular plan remains active on the device(s) you intend to keep and use, please reply back within that time frame with any information you might have."
    $body += "<br />"
    $body += "<br />"
    $body += "Thank you!"


    Send-MailMessage -SmtpServer "mailserver.example.com" -from $from -to $recipients -Subject $subject -Body $body -BodyAsHtml

}










# Determines path of the script
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent

# Gets all the csv files in the folder
$ImportFiles = Import-Csv -Path (Get-ChildItem -Path "$path\" -Recurse -Filter '*.csv').FullName | Where 'Phone Number' -NotMatch 'Report Created At'

# This is an array that will hold all OOC entries
$OOCArray = @()

# We go through the files and create a custom object that holds all the attribute (because the ordering for Device Custom Attribute isn't consistent)
foreach ($item in $ImportFiles){

    # Placeholder attributes
    $deviceRole = "N/A"
    $deviceLocation = "N/A"
    
    # Because the .csv export from MobileIron is not consistent when it sends the custom attributes, we need to format our data to ensure we process it correctly
    # We split the customAttribute value
    $customAttribute0 = (($item.'Device Custom Attributes' -split ',')[0] -split ':')[0]
    $customAttribute1 = (($item.'Device Custom Attributes' -split ',')[1] -split ':')[0]

    # We clean the attributes up
    $customAttribute0 = $customAttribute0 -replace '[[\]{}''"]'
    $customAttribute1 = $customAttribute1 -replace '[[\]{}''"]'
    
    # We determine which attribute is which
    Switch($customAttribute0){
        "devicerole"{
            $deviceRole = (($item.'Device Custom Attributes' -split ',')[0] -split ':')[1]
        }
        "devicelocation"{
            $devicelocation = (($item.'Device Custom Attributes' -split ',')[0] -split ':')[1]
        }
    }
    Switch($customAttribute1){
        "devicerole"{
            $deviceRole = (($item.'Device Custom Attributes' -split ',')[1] -split ':')[1]
        }
        "devicelocation"{
            $devicelocation = (($item.'Device Custom Attributes' -split ',')[1] -split ':')[1]
        }
    }

    # More cleaning up
    $deviceRole = $deviceRole -replace '[[\]{}''"]'
    $deviceLocation = $deviceLocation -replace '[[\]{}''"]'


    # Now that we know the info, we can create a custom PowerShell obj to make things easier / readable
    $obj = New-Object -Type PSObject -Property @{
        'DeviceNumber' = $item.'Phone Number'
        'DeviceSerial' = $item.'Serial Number'
        'DeviceModel' = $item.Model
        'DeviceCarrier' = $item.'Current Carrier Network'
        'DeviceRole'   = $deviceRole
        'DeviceLocation' = $deviceLocation
    }
    
    # We add that obj to the array (that will hold all objects)
    $OOCArray += $obj

}


# Now that the OOC array is populated, we sort by property so we can ensure that we send a property a list of all OOC devices at their property
$OOCArray = $OOCArray | Sort-Object -Property 'DeviceLocation'


# Placeholder to use for logic stuff
$placeholderProperty = ""
$placeholderItems = @()

foreach($device in $OOCArray){

    # If there is an existing serial number (i.e. filters out the last row or any weird entries that doesn't have a serial number we can work with)
    If ($device.DeviceSerial){

        # We use a placeholder to determine if an item exists to add the devices to an array or if a new array needs to be made for the next property

        # If the placeholder property exists in the entry (i.e. this isn't the very first run)
        If ($placeholderProperty){

            # We add to the existing list of devices at the placeholder property
            If($placeholderProperty.Equals($device.DeviceLocation)){
            
                # Expand and add the current device to the list
                $placeholderItems = $placeholderItems + $device
            }
            Else{
                # We are done with the property. We need to send the email, then create a new list with the new device entry

                # Send email by calling function
                Send-OOC-Email -placeholderItems $placeholderItems -placeholderProperty $placeholderProperty


                # Now we update the current placeholder property to the new property of this device
                $placeholderProperty = $device.DeviceLocation
                # Clear the array and variables
                $placeholderItems = @()
                # Add the device to the list of the placeholder devices
                $placeholderItems = $placeholderItems + $device
            }

        }
        # PlaceholderProperty doesn't exist; this is the first run
        else {
            $placeholderProperty = $device.DeviceLocation
            $placeholderItems = $placeholderItems + $device
        }


    }
}












