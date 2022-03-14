

####################################
# Author:       Eric Austin
# Create date:  
# Description:  
####################################

#namespaces
using namespace System.Data;    #required for DataTable
using namespace System.Data.SqlClient;
using namespace System.Collections.Generic; #required for List<T>
using module Send-MailKitMessage;

Try {

    #common variables
    $CurrentDirectory = [string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot;  #$PSScriptRoot is an empty string when not run from a script, and null coalescing doens't work with empty strings
    $CurrentDirectory = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {(Get-Location).Path} else {$PSScriptRoot}; #for a Windows PowerShell script
    $ErrorActionPreference = "Stop";
    $ErrorData = [List[object]]::new();
    $ErrorLogLocation = Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv";

    #SQL Server parameters (may also be populated from an external sources for security)
    $Connection = [SqlConnection]::new();
    $Connection.ConnectionString = "Server=ServerName;Database=DatabaseName;User Id=Username;Password=Password;";
    $Connection.ConnectionString = "Server=ServerName;Database=DatabaseName;Trusted_Connection=True;";
    $Command = $Connection.CreateCommand();
    $Command.CommandType = [CommandType]::StoredProcedure;  #if command will be a proc
    $Command.CommandTimeout = 1000; #seconds to wait for command to execute, default is 30 seconds

    #initialize SQL Server SqlBulkCopy object and parameters
    $Bulk = [SqlBulkCopy]::new($Connection);
    $Bulk.BatchSize = 1000; #default batch size is 1
    $Bulk.DestinationTableName = "";

    #initialize datatable for use with SqlBulkCopy
    $Datatable = [DataTable]::new();
    $Datatable.Columns.Add("RowCreateDate", [datetime]) | Out-Null;
    $Datatable.Columns.Add("OrderYear", [int]) | Out-Null;
    $Datatable.Columns.Add("OrderMonth", [string]) | Out-Null;
    
    #disallow null values
    $Datatable.Columns["RowCreateDate"].AllowDBNull = $false | Out-Null;
    $Datatable.Columns["OrderYear"].AllowDBNull = $false | Out-Null;
    $Datatable.Columns["OrderMonth"].AllowDBNull = $false | Out-Null;
    
    #add a primary key
    $Datatable.PrimaryKey = [DataColumn[]]($Datatable.Columns["OrderYear"], $Datatable.Columns["OrderMonth"]) | Out-Null;
    
    #add a unique constraint
    $Datatable.Constraints.Add([UniqueConstraint]::new([DataColumn[]]($Datatable.Columns["OrderYear"], $Datatable.Columns["OrderMonth"]))) | Out-Null;

    #map SqlBulkCopy datatable columns to SQL table columns
    $Bulk.ColumnMappings.Add("RowCreateDate", "RowCreateDate") | Out-Null;
    $Bulk.ColumnMappings.Add("OrderYear", "OrderYear") | Out-Null;
    $Bulk.ColumnMappings.Add("OrderMonth", "OrderMonth") | Out-Null;

    #script elements
    $FilesToImportDirectory = null;
    $FilesAlreadyImportedDirectory = null;
    $Date = (Get-Date).ToString();  #returns "6/20/2019 9:10:21 AM" for use in log entries
    $Date = (Get-Date).ToString("yyyyMMdd");    #returns "20190620" for use in file or folder names
    $Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); #returns "2019-06-20 09:16:41" for use in SQL Server
    $List1 = [List[object]]::new(); #list of type object which can be filled with PSCustomObjects

    #--------------#
    
    #ensure the ImportExcel module is installed
    if ( -not (Get-Module -ListAvailable | Where-Object -Property Name -EQ "ImportExcel"))
    {
        Throw "The ImportExcel module (Install-Module ImportExcel -Scope CurrentUser) is required for importing. Processing cancelled.";
    }
    
    #ensure "Files to import" directory exists
    if (-not (Test-Path $FilesToImportDirectory))
    {
        New-Item -ItemType Directory -Path $FilesToImportDirectory | Out-Null;
    }
    
    #import files and move them
    foreach ($File in $FilesToImportDirectory){
        Move-Item -Path $File."FullName" -Destination $FilesAlreadyImportedDirectory;
    }
    
    $Connection.Open();

    #SQL Server - individual insertion
    $Command.CommandText = "INSERT INTO TaskSchedulerLog (TaskName, NextRunTime) VALUES (@TaskName, @NextRunTime)";
    $Command.Parameters.Add("@TaskName", [SqlDbType]::VarChar,1000).Value = $_.TaskName;
    $Command.Parameters.Add("@NextRunTime", [SqlDbType]::DateTime).Value = if (($_."Next Run Time" -eq "N/A") -or ($_."Next Run Time" -eq "11/30/1999 12:00:00 AM")) {[System.DBNull]::Value} else {$_."Next Run Time"};
    $Command.ExecuteNonQuery() | Out-Null;

    #SQL Server - bulk insertion
    $Datatable.Rows.Add("2018/02/01", "2018", "January");
    $Datatable.Rows.Add("2019/07/01", "2019", "June");
    $Bulk.WriteToServer($Datatable);

    #SQL Server - return resultset from command
    $Datatable.Dispose();
    $Datatable = [DataTable]::new();
    $Command.CommandText = "SELECT TOP 10 ID, Color FROM LogData;";
    $Adapter = [SqlDataAdapter]::new($Command);
    $Adapter.Fill($Datatable) | Out-Null;
    
    #SQL Server - return resultset from proc
    $Command.CommandText = "dbo.ProcName";
    $Adapter = [SqlDataAdapter]::new($Command);
    $Adapter.Fill($Datatable) | Out-Null;
    
    #export datatable to Excel (exclude properties which get included as columns when exporting a SQL result)
    $Datatable | Export-Excel -Path $FileDestinationLocation -ExcludeProperty "RowError", "RowState", "Table", "ItemArray", "HasErrors";
    
    #SQL Server - return multiple resultsets from proc
    $Command.CommandText = "dbo.ProcName";
    $Dataset = [DataSet]::new();
    $Adapter = [SqlDataAdapter]::new($Command);
    $Adapter.Fill($Dataset) | Out-Null;
    $Dataset.Tables[0];
    $Dataset.Tables[1];
    
    $Connection.Close();

    #set email parameters
    #use secure connection if available
    $UseSecureConnectionIfAvailable = $true;

    #authentication
    $Credential = [System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force));

    #SMTP server
    $SMTPServer = "SMTPServer";

    #port
    $Port = PortNumber;

    #sender
    $From = [MimeKit.MailboxAddress]"SenderEmailAddress";

    #recipient list
    $RecipientList = [MimeKit.InternetAddressList]::new();
    $RecipientList.Add([MimeKit.InternetAddress]"Recipient1EmailAddress");

    #cc list
    $CCList = [MimeKit.InternetAddressList]::new();
    $CCList.Add([MimeKit.InternetAddress]"CCRecipient1EmailAddress");

    #bcc list
    $BCCList = [MimeKit.InternetAddressList]::new();
    $BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress");

    #subject
    $Subject = [string]"Subject";

    #text body
    $TextBody = [string]"TextBody";

    #HTML body
    $HTMLBody = [string]"HTMLBody";

    #attachment list
    $AttachmentList = [System.Collections.Generic.List[string]]::new();
    $AttachmentList.Add("Attachment1FilePath");

    #splat parameters
    $Parameters = @{
        "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable
        "Credential" = $Credential
        "SMTPServer" = $SMTPServer
        "Port" = $Port
        "From" = $From
        "RecipientList" = $RecipientList
        "CCList" = $CCList
        "BCCList" = $BCCList
        "Subject" = $Subject
        "TextBody" = $TextBody
        "HTMLBody" = $HTMLBody
        "AttachmentList" = $AttachmentList
    };

    #send message
    Send-MailKitMessage @Parameters;

    #fill a list with custom objects
    Get-Process | ForEach-Object {
        $ProcessesReport.Add(
            [PSCustomObject]@{
                ID = $_.Id
                Process = $_.ProcessName
            }
        );
    }

    #bearer API header
    $Header = @{Authorization = "Bearer PersonalAccessToken"};

    #basic API header (requires a base-64 encoded string of a username:password)
    $Header = @{ Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("Username:Token")) };

    #REST API call - Get
    Invoke-RestMethod -Method Get -Uri "https://api.github.com/user" -Headers $Header;

    #REST API call - Patch
    $Body = @{
        "twitter_username" = "@Test1"
    } | ConvertTo-Json;
    Invoke-RestMethod -Method "Patch" -Uri "https://api.github.com/user" -Headers $Header -Body $Body;

}

Catch {

    #error log
    $ErrorData.Add([PSCustomObject]@{
        Date = (Get-Date).ToString()
        ErrorMessage = $Error[0]
    });
    $ErrorData | Select-Object "Date", "ErrorMessage" | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation;

    #return value
    Exit 1;
    
}

Finally {

    if ($Connection.State -ne "Closed")
    {
        $Connection.Close();
    }

}
