$queryResults = Get-ADComputer -Filter * -Properties whenCreated, OperatingSystem, operatingSystemVersion, LastLogonDate | Select name,@{n="owner";e={(Get-acl "ad:\$($_.distinguishedname)").owner}}, whenCreated, distinguishedName, OperatingSystem, operatingSystemVersion, LastLogonDate

$fileName = 'AllComputersInAD.csv'
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$queryResults | Export-Csv -Path $path\$fileName -NoTypeInformation
