# This script looks up a user based on employee number in AD

$employeeNumber = "XXXXXXXX" #Employee number of user
Get-ADUser -Filter {EmployeeNumber -eq "009085"} -Properties DisplayName
