$filename = "kill-list.csv"

$path = $MyInvocation.MyCommand.Path | Split-Path -Parent
$file = "$path\$filename"

$csvExport = Import-CSV $file

foreach($entry in $csvExport){
     Try {
         Remove-ADComputer $entry.Name -Confirm:$False -Verbose -ErrorAction STOP
         "Removed $entry.Name" | Out-File -FilePath "$path\RemovedComputers.txt" -Append
     }
     Catch{
         "Failed to removed $entry.Name" | Out-File -FilePath "$path\FailedComputers.txt" -Append
         $_ | Out-File -FilePath "$path\RemovedComputers.txt" -Append     # writes the exception
    }

}
