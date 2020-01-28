
####################################
# Author:       Eric Austin
# Description:  This is a script template. It returns an exit code of 99 on error and creates an error log in the directory it is run out of.
####################################

#declare namespaces
using namespace System.Data
using namespace System.Data.SqlClient
using namespace System.Data.SQLite

#set common variables
$CurrentDirectory=if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
$ErrorActionPreference="Stop"
$ErrorData=@()
$ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"

#files to import
$FilesToImportDirectory=""
$AlreadyImportedDirectory=""

#dot source function
. "$CurrentDirectory\FunctionTemplate.ps1"

#add SQLite assembly
Add-Type -Path "$CurrentDirectory\System.Data.SQLite\System.Data.SQLite.dll"

#common date variables
$Date=(Get-Date).ToString() #returns "6/20/2019 9:10:21 AM" for use in log entries
$Date=(Get-Date).ToString("yyyyMMdd") #returns "20190620" for use in file or folder names
$Date=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") #returns "2019-06-20 09:16:41" for use in SQL Server

#SQL Server parameters (may also be populated from an external file for security)
#use the .Net method of interacting with SQL Server even though it's more verbose because it's the safe way to avoid SQL injection
$Server=""
$Database=""
$Username=""
$Password=''
$Connection=New-Object SqlConnection
$ConnectionString="Server=$Server;Database=$Database;User Id=$Username;Password=$Password;"
$ConnectionString="Server=$Server;Database=$Database;Trusted_Connection=True;"
$Connection.ConnectionString=$ConnectionString
$Cmd=$Connection.CreateCommand()
$Cmd.CommandTimeout=1000 #number of time in seconds to wait for the command to execute, default is 30 seconds

#initialize SQL Server SqlBulkCopy object and parameters
$Bulk=New-Object SqlBulkCopy($Connection)
$Bulk.BatchSize=10000   #default batch size is 1
$Bulk.DestinationTableName="Orders"

#initialize datatable for use with SqlBulkCopy
$Datatable=New-Object DataTable
$Datatable.Columns.Add("RowCreateDate", "Datetime") | Out-Null
$Datatable.Columns.Add("OrderYear", "Int16") | Out-Null
$Datatable.Columns.Add("OrderMonth", "String") | Out-Null

#map SqlBulkCopy datatable columns to SQL table columns
$Bulk.ColumnMappings.Add("RowCreateDate", "RowCreateDate") | Out-Null
$Bulk.ColumnMappings.Add("OrderYear", "OrderYear") | Out-Null
$Bulk.ColumnMappings.Add("OrderMonth", "OrderMonth") | Out-Null

#SQLite parameters
$SLConnection=New-Object SQLiteConnection
$SLConnectionString="Data Source=$CurrentDirectory\SQLite.db"
$SLConnection.ConnectionString=$SLConnectionString
$SLCmd=$SLConnection.CreateCommand()

#initialize array for function splatting
$DataForFunction=@()

Try {
    
    $Connection.Open()

    #import files
    Get-ChildItem -Path $FilesToImportDirectory -Include *.csv | ForEach-Object {

        Move-Item -Path $_.FullName -Destination $AlreadyImportedDirectory

    }

    #SQL Server - individual insertion
    $Cmd.CommandText="INSERT INTO TaskSchedulerLog (TaskName, NextRunTime) VALUES (@TaskName, @NextRunTime)"
    $Cmd.Parameters.Add("@TaskName", [SqlDbType]::VarChar,1000).Value=$_.TaskName
    $Cmd.Parameters.Add("@NextRunTime", [SqlDbType]::DateTime).Value=if (($_."Next Run Time" -eq "N/A") -or ($_."Next Run Time" -eq "11/30/1999 12:00:00 AM")) {[System.DBNull]::Value} else {$_."Next Run Time"}
    $Cmd.ExecuteNonQuery() | Out-Null

    #SQL Server - bulk insertion
    $Datatable.Rows.Add("2018/02/01", "2018", "January")
    $Datatable.Rows.Add("2019/07/01", "2019", "June")
    $Bulk.WriteToServer($Datatable)

    #SQL Server - return resultset
    $Datatable.Dispose()
    $Datatable=New-Object DataTable
    $Cmd.CommandText="SELECT TOP 10 ID, Color FROM LogData;"
    $Adapter=New-Object SqlDataAdapter $Cmd
    $Adapter.Fill($Datatable) | Out-Null

    #SQLite - import pipe-delimited file using command line utility
    $Import="sqlite3 Sync.db -cmd "".mode csv"" -cmd "".separator |"" -cmd "".import 'FileToImport' TableToImportTo"" "".exit"" 2>&1"
    Invoke-Expression $Import

    #SQLite - execute query
    $SLCmd.CommandText="INSERT INTO Table1 (ID) SELECT ID FROM Table2;"
    $SLCmd.ExecuteNonQuery() | Out-Null

    #SQLite - return resultset
    $Datatable.Dispose()
    $Datatable=New-Object DataTable
    $SLCmd.CommandText="SELECT TOP 10 ID, Color FROM LogData;"
    $SLAdapter=New-Object SQLiteDataAdapter $SLCmd
    $SLAdapter.Fill($Datatable) | Out-Null

    #set splatting values and call function
    $DataForFunction=@{
        "ID"="789"
        "Color"="Red"
    }
    ExampleFunction @DataForFunction
}

Catch {

    #error log
    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{"Date"=(Get-Date).ToString(); "ErrorMessage"=$Error[0].ToString()}
    $ErrorData | Select-Object Date,ErrorMessage | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation

    #return value
    Exit 99
    
}

Finally {

    $Connection.Close()

}
