
####################################
# Author:       Eric Austin
# Create date:  
# Description:  This is a script template. It returns an exit code of 99 on error and creates an error log in the directory it is run out of.
####################################

#namespaces
using namespace System.Data     #required for DataTable
using namespace System.Data.SqlClient   
using namespace System.Data.SQLite

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"
    $ErrorLogLocation=(Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv")

    #files to import
    $FilesToImportDirectory=""
    $AlreadyImportedDirectory=""

    #dot source function
    . "$CurrentDirectory\FunctionTemplate.ps1"

    #add SQLite assembly
    Add-Type -Path "$CurrentDirectory\System.Data.SQLite\System.Data.SQLite.dll"

    #date variables
    $Date=(Get-Date).ToString() #returns "6/20/2019 9:10:21 AM" for use in log entries
    $Date=(Get-Date).ToString("yyyyMMdd") #returns "20190620" for use in file or folder names
    $Date=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") #returns "2019-06-20 09:16:41" for use in SQL Server

    #SQL Server parameters (may also be populated from an external sources for security)
    $Connection=New-Object SqlConnection
    $Connection.ConnectionString="Server=ServerName;Database=DatabaseName;User Id=Username;Password=Password;"
    $Connection.ConnectionString="Server=ServerName;Database=DatabaseName;Trusted_Connection=True;"
    $Command=$Connection.CreateCommand()
    $Command.CommandType=[CommandType]::StoredProcedure #if command will be a proc
    $Command.CommandTimeout=1000    #number of time in seconds to wait for the command to execute, default is 30 seconds

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

    #--------------#
    
    #ensure the ImportExcel module is installed
    if ( -not (Get-Module -ListAvailable | Where-Object -Property Name -EQ "ImportExcel"))
    {
        Throw "The ImportExcel module (Install-Module ImportExcel -scope CurrentUser) is required for importing. Processing aborted."
    }
    
    #import files
    Get-ChildItem -Path $FilesToImportDirectory -Include *.csv | ForEach-Object {

        Move-Item -Path $_.FullName -Destination $AlreadyImportedDirectory

    }
    
    $Connection.Open()

    #SQL Server - individual insertion
    $Command.CommandText="INSERT INTO TaskSchedulerLog (TaskName, NextRunTime) VALUES (@TaskName, @NextRunTime)"
    $Command.Parameters.Add("@TaskName", [SqlDbType]::VarChar,1000).Value=$_.TaskName
    $Command.Parameters.Add("@NextRunTime", [SqlDbType]::DateTime).Value=if (($_."Next Run Time" -eq "N/A") -or ($_."Next Run Time" -eq "11/30/1999 12:00:00 AM")) {[System.DBNull]::Value} else {$_."Next Run Time"}
    $Command.ExecuteNonQuery() | Out-Null

    #SQL Server - bulk insertion
    $Datatable.Rows.Add("2018/02/01", "2018", "January")
    $Datatable.Rows.Add("2019/07/01", "2019", "June")
    $Bulk.WriteToServer($Datatable)

    #SQL Server - return resultset from command
    $Datatable.Dispose()
    $Datatable=New-Object DataTable
    $Command.CommandText="SELECT TOP 10 ID, Color FROM LogData;"
    $Adapter=New-Object SqlDataAdapter $Command
    $Adapter.Fill($Datatable) | Out-Null
    
    #SQL Server - return resultset from proc
    $Command.CommandText="dbo.ProcName"
    $Adapter=New-Object SqlDataAdapter $Command
    $Adapter.Fill($Datatable) | Out-Null
    
    #SQL Server - return multiple resultsets from proc
    $Command.CommandText="dbo.ProcName"
    $Dataset=New-Object DataSet
    $Adapter=New-Object SqlDataAdapter $Command
    $Adapter.Fill($Dataset) | Out-Null
    $Dataset.Tables[0]
    $Dataset.Tables[1]

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
    Exit 1
    
}

Finally {

    $Connection.Close()

}
