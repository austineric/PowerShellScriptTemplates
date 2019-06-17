
####################################
# Author:      Eric Austin
# Create date: June 2019
# Description: This is a script template. It returns an exit code of -99 for use with Task Scheduler (to be used with Exec_PSScript.ps1) and creates an error log in the directory it is run out of.
####################################

$CurrentDirectory= if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
$ErrorActionPreference="Stop"
$ErrorData=@()
$ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"

Try {
    

}
Catch {

    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{"Date"=(Get-Date).ToString(); "ErrorMessage"=$Error[0].ToString()}
    $ErrorData | Select-Object Date,ErrorMessage | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation
    Exit -99
    
}
