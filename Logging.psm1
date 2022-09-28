<#
.SYNOPSIS
  Logging helper scripts

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Created by Paul Shiryaev <ps@paulshiryaev.com> 
  github.com/ps5

  How to import this module:
  Import-Module -Name D:\Tasks\PS\Logging.psm1 -Force -ArgumentList 'Persist Security Info=False;Integrated Security=true;Initial Catalog=ADMIN;server="$YourAdminSqlServerName"'

  $YourAdminSqlServerName is your SQL Server which contains administrative (ADMIN) database

#>

param(
    [parameter(Position=0,Mandatory=$true)][string]$TargetCnnStr
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  # disable Tls11

[bool] $Global:DebugMode = $false # set to $true to prevent logging to the production log
[string] $Global:DatabaseName = "ADMIN" # override name of the administrative db here
[string] $Global:SPAuditProc = $Global:DatabaseName+".logs.sp_audit_proc"

# Import-Module -Name "$PSScriptRoot\ETL.psm1" -Force
# Import-Module -Name "$PSScriptRoot\TSQL.psm1" -Force
  
function Write-AuditLog {
param(
  [string] $LogState
  ,[string] $LogMessage
  ,[string] $ProcName = $null
  # ,[string] $ServerName = $Global:ServerName
  ,[string] $RowCount = $null
  ,[boolean] $NoOutput = $False
  )

    if ($ProcName -eq $null -or $ProcName -eq "") {

        $ProcName = $MyInvocation.MyCommand.Name
        $callStack = Get-PSCallStack
        if ($callStack.Count -gt 1) {
            # $ProcName  = $callStack[1].FunctionName
            $ProcName  = $callStack[1].ScriptName
        }
        $ProcName = "["+$env:computername+"] "+$ProcName
    }

  try {

        $now = $(Get-Date).ToString()
	if (!$NoOutput) {
	        Write-Output  "$now $ProcName $LogState $LogMessage"
	}

        if (!$DebugMode) {

            $conn = New-Object System.Data.SqlClient.SqlConnection
            $conn.ConnectionString = $TargetCnnStr # "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=LOGGING.PS"

            $cmd = $conn.CreateCommand();
            $cmd.CommandText = $Global:SPAuditProc
            $cmd.CommandType = [System.Data.CommandType]::StoredProcedure
            $cmd.Parameters.AddWithValue("@ProcName", $ProcName) | Out-Null
            $cmd.Parameters.AddWithValue("@State", $LogState) | Out-Null
            $cmd.Parameters.AddWithValue("@LogMessage", $LogMessage) | Out-Null
            $cmd.Parameters.AddWithValue("@rowcount", $RowCount) | Out-Null

            $conn.open()
            $rc = $cmd.ExecuteNonQuery();
            $conn.Close()

        }

   }

   catch {
       $criticalerror = 1
       Write-Error $Error[0].Exception.Message
       # break
   }

}




function Get-LastLogDate {
param(
       [string]$connStr
       ,[string]$key
       )

    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    $cmd = $conn.CreateCommand();
    $cmd.CommandText = "SELECT TOP 1 [Value] FROM ADMIN.[logs].[Tracker] WHERE [Key] = @key;"
    $cmd.Parameters.AddWithValue("@key", $key) | Out-Null

    $dataTable = New-Object System.Data.DataTable
    $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $sqlAdapter.SelectCommand = $cmd
    $rc = $sqlAdapter.Fill($dataTable)

    foreach ($dataRow in $dataTable.Rows)
    {    
        $CurrentLogDate = $dataRow["Value"].ToString()
    }

    $conn.Close()
    return $CurrentLogDate
}


function Set-LastLogDate {
param(
       [string]$connStr
       ,[string]$key
       ,[string]$NewLogDate
       )

    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    # Write-Host $NewLogDate
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "update ADMIN.logs.tracker set [value] = @value where [key] = @key"
    $cmd.Parameters.AddWithValue("@value", $NewLogDate) | Out-Null
    $cmd.Parameters.AddWithValue("@key", $key) | Out-Null
    $rc = $cmd.ExecuteNonQuery()
    if ($rc -eq 0) { # insert
        $cmd.CommandText = "insert into ADMIN.logs.tracker ([value], [key]) values (@value, @key)"
        $rc = $cmd.ExecuteNonQuery()
    }
    $conn.Close()
}





function Get-EventLogs {
param ($ComputerName)

    # Get recent shutdowns
    Get-Eventlog -LogName System -Source "User32" -ComputerName $ComputerName | Out-GridView

    # Get recent RDP connections
    Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{LogName='Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'; Id='1149' } | fl

}





function pushMessageCardJson {
param(
    [string]$Json
    ,[string]$Uri
       )

    # Write-Host $Json

    $Response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Method POST -Body @($Json) 
    # Write-Host $Response 

}

function Push-MSTeamsAlerts {
    param(
       [string]$ServerName=$(throw 'server name is required.'),
       [string]$Key="MSTeams",
       [string]$MSTeamsAlertsDbView = "ADMIN.[logs].[vw_MSTeamsAlerts]"
       )

    # INIT
    $criticalerror = 0
    $verbose = 0
       
    # Test connect
    try {
        if ($verbose) { Write-Host $(Get-Date) "Connecting to $ServerName..." }
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName"
        $conn.open()
                
        $LogDate = Get-LastLogDate -connStr $conn.ConnectionString -key $Key

        $cmd = $conn.CreateCommand();
        $cmd.CommandText = "SELECT LogDate, Json, Uri FROM $MSTeamsAlertsDbView WHERE dateadd(second, -1, LogDate) > '$LogDate' ORDER BY LogDate ASC;"

        $dataTable = New-Object System.Data.DataTable
        $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

        $sqlAdapter.SelectCommand = $cmd
        $rc = $sqlAdapter.Fill($dataTable)

        foreach ($dataRow in $dataTable.Rows)
        {   
            try { 
                $CurrentLogDate = $dataRow["LogDate"].ToString()
                $Json = $dataRow["Json"]
                $Uri = $dataRow["Uri"]
                Write-Host $(Get-Date) $CurrentLogDate
    
                # MS Teams
                pushMessageCardJson -Json $Json -Uri $Uri  ## "https://outlook.office.com/webhook/../IncomingWebhook/.."
            
                Set-LastLogDate -connStr $conn.ConnectionString -key $key -NewLogDate $CurrentLogDate
            }
            catch {
                # Write-Host $Json
                Write-AuditLog -LogState 'E' -LogMessage $Error[0].Exception.Message
                # Write-Host $Error[0].ToString()
            }
        }


        $conn.Close()
    }

    catch {
        $criticalerror = 1
        Write-AuditLog -LogState 'E' -LogMessage $Error[0].Exception.Message
        break
    }

}






Export-ModuleMember *-*

