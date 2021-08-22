$sAMAccountName = "RTest"


function Separate($sAMAccountName){

      # Remove any membership
      $userInfo = Get-ADUser -Identity $sAMAccountName -Properties MemberOf 
      ForEach ($group in $userInfo.MemberOf)
      {
          Remove-ADGroupMember -Identity $group -Members $sAMAccountName -Confirm:$false
      }

      # Hide in Address Book
      Set-ADUser -identity $sAMAccountName -Replace @{msExchHideFromAddressLists = $true}

      # Disable account
      Disable-ADAccount -identity $sAMAccountName

      # Clear manager
      Set-ADUser -identity $sAMAccountName -manager $null

      # Move to separated OU
      Get-ADUser $sAMAccountName | Move-ADObject -TargetPath 'OU=Separated,DC=corp,DC=example, DC=com'
    
}
