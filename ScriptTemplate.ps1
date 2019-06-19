
####################################
# Author:      Eric Austin
# Create date: June 2019
# Description: This is a script template. It returns an exit code of -99 for use with Task Scheduler (to be used with Exec_PSScript.ps1) and creates an error log in the directory it is run out of.
####################################

using namespace System.Data.SqlClient

$CurrentDirectory=if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
$ErrorActionPreference="Stop"
$ErrorData=@()
$ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"

#SQL Server parameters if needed (may also be populated from an external file for security)
#use the .Net method of interacting with SQL Server even though it's more verbose because it's the only safe way to avoid SQL injection
$Server=""
$Database=""
$Username=""
$Password=''
$Connection=New-Object System.Data.SqlClient.SqlConnection
$ConnectionString="Server=$Server;Database=$Database;User Id=$Username;Password=$Password"
$Connection.ConnectionString=$ConnectionString

Try {
    
    $Connection.Open()
    $Cmd=$Connection.CreateCommand()
    $Cmd.CommandText="INSERT INTO VAPurchaseOrders_Staging (XMLText) VALUES (@value)"
    $Cmd.CommandTimeout=120
    $Cmd.Parameters.Add("@value", [System.Data.SqlDbType]::Varchar, '-1').Value=$Data
    $Cmd.ExecuteNonQuery() | Out-Null

}

Catch {

    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{"Date"=(Get-Date).ToString(); "ErrorMessage"=$Error[0].ToString()}
    $ErrorData | Select-Object Date,ErrorMessage | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation
    Exit -99
    
}

Finally {

    $Connection.Close()

}
