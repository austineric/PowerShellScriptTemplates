
####################################
# Author:      Eric Austin
# Create date: June 2019
# Description: Script template demonstrating SQLite usage
####################################

using namespace System.Data
using namespace System.Data.SQLite

$CurrentDirectory=if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
Add-Type -Path "$CurrentDirectory\SystemDataSQLite\System.Data.SQLite.dll"

#SQL Server connection parameters required
$Connection=New-Object SQLiteConnection
$Connection.ConnectionString="Data Source=$CurrentDirectory\SQLite.db"
$Cmd=$Connection.CreateCommand()
$Cmd.CommandText="SELECT date() AS 'Date' UNION ALL SELECT date();"
$Adapter=New-Object SQLiteDataAdapter $Cmd
$Dt=New-Object Datatable
$Connection.Open()
$Adapter.Fill($Dt)
$Connection.Close()


 
$Dt.Date[0].GetType() 





#build datatable
$Datatable=New-Object DataTable
$Datatable.Columns.Add("RowCreateDate", "Datetime") | Out-Null
$Datatable.Columns.Add("OrderYear", "Int16") | Out-Null
$Datatable.Columns.Add("OrderMonth", "String") | Out-Null

#fill datatable
$Datatable.Rows.Add($RowCreateDate, $OrderYear, $OrderMonth)

#create SqlBulkCopy object and parameters
$Bulk=New-Object SqlBulkCopy($Connection)
$Bulk.BatchSize=10000   #default batch size is 1
$Bulk.DestinationTableName="Orders"

#SqlBulkCopy just lines datatable columns up to the SQL table columns so having for example an identity column in the first position will break the insert unless column mappings are created
$Bulk.ColumnMappings.Add("RowCreateDate, RowCreateDate") | Out-Null
$Bulk.ColumnMappings.Add("OrderYear, OrderYear") | Out-Null
$Bulk.ColumnMappings.Add("OrderMonth, OrderMonth") | Out-Null

#write data to SQL Server
$Connection.Open()
$Bulk.WriteToServer($Datatable)
$Connection.Close()
