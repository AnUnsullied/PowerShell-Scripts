$filename = "Adobe-Download.csv"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$file = "$path\$filename"

$csvExport = Import-CSV $file | Where 'Product Configurations' -ne ""

foreach ($item in $csvExport){
    $email = $item.Email
    $user = Get-ADUser -Filter {EmailAddress -eq $email}


    if(-not ($user.Enabled)){



        $body = "This PowerShell script found some issues with some accounts in Adobe. Please review the account in Adobe and address accordingly:`n"
        $body += ("`nAdobe Account Email: " + $item.Email +"`n" )
        $body += ("`nAdobe Licenses Assigned: " + $item.'Product Configurations')

        $body += "`n"

        if($user){
            $body += ("`nEnabled in AD?: "+ $user.Enabled)
            $body += ("`nAD Account Location: " + $user)
        }
        else{
            $body += "`nNo AD Account Found."
        }
         
        # Send email to the IT helpdesk to create tickets for users that are disabled / terminated in Active Directory but still has an Adobe license.
        Send-MailMessage -From "self-email@example.com" -To "IT-helpdesk@example.com" -SmtpServer "server.example.com" -Subject "John's Adobe Licensing Script Check" -Body $body
    }
}
