
####################################
# Author:      Eric Austin
# Create date: June 2019
# Description: Script template demonstrating .Net's SqlBulkCopy
####################################

using namespace System.Data
using namespace System.Data.SqlClient

#SQL Server connection parameters required
$Connection=New-Object SqlConnection

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
