
####################################
# Author:      Eric Austin
# Create date: June 2019
# Description: This is a script template. It returns an exit code of -99 for use with Task Scheduler (to be used with Exec_PSScript.ps1) and creates an error log in the directory it is run out of.
####################################

using namespace System.Data
using namespace System.Data.SqlClient

$CurrentDirectory=if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
$ErrorActionPreference="Stop"
$ErrorData=@()
$ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"

$Date=(Get-Date).ToString() #returns "6/20/2019 9:10:21 AM" for use in log entries
$Date=(Get-Date).ToString("yyyyMMdd") #returns "20190620" for use in file or folder names
$Date=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") #returns "2019-06-20 09:16:41" for use in SQL Server

#SQL Server parameters if needed (may also be populated from an external file for security)
#use the .Net method of interacting with SQL Server even though it's more verbose because it's the only safe way to avoid SQL injection
$Server=""
$Database=""
$Username=""
$Password=''
$Connection=New-Object SqlConnection
$ConnectionString="Server=$Server;Database=$Database;User Id=$Username;Password=$Password;"
$ConnectionString="Server=$Server;Database=$Database;Trusted_Connection=True;"
$Connection.ConnectionString=$ConnectionString

Try {
    
    $Connection.Open()
    $Cmd=$Connection.CreateCommand()
    $Cmd.CommandText="INSERT INTO TaskSchedulerLog (TaskName, NextRunTime) VALUES (@TaskName, @NextRunTime)"
    $Cmd.Parameters.Add("@TaskName", [SqlDbType]::VarChar,1000).Value=$_.TaskName
    $Cmd.Parameters.Add("@NextRunTime", [SqlDbType]::DateTime).Value=if (($_."Next Run Time" -eq "N/A") -or ($_."Next Run Time" -eq "11/30/1999 12:00:00 AM")) {[System.DBNull]::Value} else {$_."Next Run Time"}
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
