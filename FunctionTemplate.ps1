

####################################
# Author:       Eric Austin
# Create date:  
# Description:  
####################################


function ExampleFunction
    (
    $ID
    ,$Color
    )

{
    $CurrentDirectory=if ($PSScriptRoot -ne "") {$PSScriptRoot} else {(Get-Location).Path}
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation="$CurrentDirectory\ErrorLog.csv"  

    $Date=(Get-Date).ToString() #returns "6/20/2019 9:10:21 AM" for use in log entries
    $Date=(Get-Date).ToString("yyyyMMdd") #returns "20190620" for use in file or folder names
    $Date=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") #returns "2019-06-20 09:16:41" for use in SQL Server

    Try {

        Write-Host $ID, $Color

    }

    Catch {

        return -99

    }
}






