$dateQuery = [DateTime]::Today.AddDays(-30)
$queryResults = Get-ADComputer -filter 'WhenCreated -ge $dateQuery' -Properties whenCreated, OperatingSystem, operatingSystemVersion | Select name,@{n="owner";e={(Get-acl "ad:\$($_.distinguishedname)").owner}}, whenCreated, distinguishedName, OperatingSystem, operatingSystemVersion

$fileName = 'Computers_from_' + $dateQuery.ToString('yyyyMMdd') + '_to_' + (Get-Date).toString('yyyyMMdd') + '.csv'
$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$queryResults | Export-Csv -Path $path\$fileName -NoTypeInformation
