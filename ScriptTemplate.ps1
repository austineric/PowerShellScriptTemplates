

####################################
# Author:       Eric Austin
# Create date:  
# Description:  
####################################

#namespaces
using namespace System.Data     #required for DataTable
using namespace System.Data.SqlClient
using namespace System.Collections.Generic  #required for List<T>
using module Send-MailKitMessage    #module classes are not automatically loaded when referenced (as opposed to commandlets which are)

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot #$PSScriptRoot is an empty string when not run from a script, and null coalescing doens't work with empty strings
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation=Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv"

    #directories
    $FilesToImportDirectory=""

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
    $Command.CommandTimeout=1000    #seconds to wait for command to execute, default is 30 seconds

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

    #initialize array for function splatting
    $DataForFunction=@()

    #script elements
    [List[object]]$List1=New-Object -TypeName List[object]  #list of type object which can be filled with PSCustomObjects
    [List[string]]$AttachmentList=New-Object -TypeName List[string] #list of string which can be filled with the location(s) of file attachments

    #--------------#
    
    #ensure the ImportExcel module is installed
    if ( -not (Get-Module -ListAvailable | Where-Object -Property Name -EQ "ImportExcel"))
    {
        Throw "The ImportExcel module (Install-Module ImportExcel -Scope CurrentUser) is required for importing. Processing aborted."
    }
    
    #ensure the Send-MailKitMessage module is installed
    if ( -not (Get-Module -ListAvailable | Where-Object -Property Name -EQ "Send-MailKitMessage"))
    {
        Throw "The Send-MailKitMessage module (Install-Module Send-MailKitMessage -Scope CurrentUser) is required for importing. Processing aborted."
    }
    
    #ensure "Files to import" directory exists
    if (-not (Test-Path $FilesToImportDirectory))
    {
        New-Item -ItemType Directory -Path $FilesToImportDirectory | Out-Null
    }
    
    #import files and move them
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
    
    #export datatable to Excel (exclude properties which get included as columns when exporting a SQL result)
    $Datatable | Export-Excel -Path $FileDestinationLocation -ExcludeProperty "RowError", "RowState", "Table", "ItemArray", "HasErrors"
    
    #SQL Server - return multiple resultsets from proc
    $Command.CommandText="dbo.ProcName"
    $Dataset=New-Object DataSet
    $Adapter=New-Object SqlDataAdapter $Command
    $Adapter.Fill($Dataset) | Out-Null
    $Dataset.Tables[0]
    $Dataset.Tables[1]
    
    $Connection.Close()

    #set splatting values and call function
    $DataForFunction=@{
        "ID"="789"
        "Color"="Red"
    }
    ExampleFunction @DataForFunction

    #set email parameters
    #from
    $From=New-Object MailboxAddressExtended($null, "SenderEmailAddress")
    
    #to
    $ToList=New-Object InternetAddressListExtended
    $ToList.Add("Recipient1EmailAddress")
    $ToList.Add("Recipient2EmailAddress")
    
    #cc
    $CCList=New-Object InternetAddressListExtended
    $CCList.Add("CCRecipientEmailAddress")
    
    #bcc
    $BCCList=New-Object InternetAddressListExtended
    $BCCList.Add("BCCRecipientEmailAddress")
    
    #subject
    $Subject="Subject"

    #body
    $HTMLBody="<span style=`"background-color:green; color:blue;`">Test email</span>"

    #attachments
    $AttachmentList.Add($FileLocation)
        
    #send email
    Send-MailKitMessage -SMTPServer "SMTPServerAddress" -Port PortNumber -From $From -ToList $ToList -CCList $CCList -BCCList $BCCList -Subject $Subject -HTMLBody $HTMLBody -AttachmentList $AttachmentList

    #fill a list with custom objects
    Get-Process | ForEach-Object {
        $ProcessesReport.Add(
            [PSCustomObject]@{
                ID = $_.Id
                Process=$_.ProcessName
            }
        )
    }

}

Catch {

    #error log
    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{"Date"=(Get-Date).ToString(); "ErrorMessage"=$Error[0].ToString()}    #don't use @Date for the date, this section needs to be completely independent so nothing can ever interfere with the error log being created
    $ErrorData | Select-Object Date,ErrorMessage | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation

    #return value
    Exit 1
    
}

Finally {

    if ($Connection.State -ne "Closed")
    {
        $Connection.Close()
    }

}
