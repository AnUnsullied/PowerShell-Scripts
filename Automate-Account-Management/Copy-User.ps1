$sAMAccountName = "UserA" #sAMAccountName of new user
$mirror = "UserB" #sAMAccountName of a person to copy attributes

$groups = (Get-ADUser -Identity $mirror -Properties MemberOf | Select-Object MemberOf).MemberOf
foreach ($group in $groups)
{
    Add-ADGroupMember $group $sAMAccountName
}

$mUser = Get-ADUser $mirror -Properties *
$user = Get-ADUser $sAMAccountName -Properties *

Set-ADUser $user -City $mUser.l -StreetAddress $mUser.streetAddress -State $mUser.State -Fax $mUser.facsimileTelephoneNumber -PostalCode $mUser.postalCode -OfficePhone $mUser.OfficePhone
