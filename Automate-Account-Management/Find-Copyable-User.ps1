# Employee number of a new hire
$newHireEmployeeNumber = 'XXXXXXX'

# We get the attributes from the new hire
$newHire = Get-ADUser -Filter {EmployeeNumber -eq $newHireEmployeeNumber} -Properties *

# Look in AD to find someone with the same office name and title so we can copy any permissions over if desired or review the accounts
Get-ADUser -Filter {physicalDeliveryOfficeName -eq $newHire.physicalDeliveryOfficeName -and title -eq $newHire.Title -and Enabled -eq "true"} -Properties saMAccountName
