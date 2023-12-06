<#
.SYNOPSIS
  Various ETL helper cmdlets
  PS DBA Toolkit

.DESCRIPTION

  Copy-AzSqlTableData
  Encrypt-File 
  Export-CSVFile
  Export-SqlTableSchemaAndSampleDataToExcel
  Get-OdbcData
  Get-OleDbData
  Get-SharepointList
  Get-SQLData
  Import-SharepointList

.NOTES
  Created by Paul Shiryaev <ps@paulshiryaev.com> 
  github.com/ps5

#>

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

Import-Module "$PSScriptRoot\TSQL.psm1" -Force


function Get-OleDbData {
param(
[string]$cnnStr
,[string]$query
,[string]$csv=""   # CSV file
,[bool]$TextQualified=$true
,[string]$ColSep=","
,[string]$Encoding="unicode"
)

# The acceptable values for the Encoding parameter are as follows:
# ascii: Uses the encoding for the ASCII (7-bit) character set.
# bigendianunicode: Encodes in UTF-16 format using the big-endian byte order.
# bigendianutf32: Encodes in UTF-32 format using the big-endian byte order.
# oem: Uses the default encoding for MS-DOS and console programs.
# unicode: Encodes in UTF-16 format using the little-endian byte order.
# utf7: Encodes in UTF-7 format.
# utf8: Encodes in UTF-8 format.
# utf8BOM: Encodes in UTF-8 format with Byte Order Mark (BOM)
# utf8NoBOM: Encodes in UTF-8 format without Byte Order Mark (BOM)
# utf32: Encodes in UTF-32 format.

    $cnn = New-Object System.Data.OleDb.OleDbConnection($cnnStr)
    $cmd = New-Object System.Data.OleDb.OleDbCommand($query)
    $cmd.Connection = $cnn;

    $Apostrophe=""
    if ($TextQualified) { $Apostrophe="""" }

    #try {
        $cnn.Open()

        if ($cnn.State -eq "Open") {
            $cmd.CommandTimeout=300
            $r = $cmd.ExecuteReader()
            $rc = 0

            # get field names
            $i=0
            $s=""
            while ($i -lt $r.FieldCount) {
                $n= $r.GetName($i)
                $s=$s+$Apostrophe+$n+$Apostrophe+$ColSep
                # $t=$r.GetFieldType($i)
                # Write-Host $n $t
                $i++
            }
            $s=$s.Substring(0,$s.Length-$ColSep.Length)
            if ($csv -ne "") {
                $s | Out-File $csv -Encoding $Encoding
            } else { Write-Host $s }


            while ($r.Read()) {   

                $rc++
                $i=0
                $s=""
                while ($i -lt $r.FieldCount) {

                    $v=[string]$r[$i]
                    $t=$r.GetFieldType($i)
                    if ($v.Contains('"') -or $v.Contains(',')) # -and ($t -eq "System.String")
                    {                          
                        $v='"'+$v.Replace('"','""')+'"' # mandatory quote strings that contain apostrophes or commas                        
                    } else { $v=$Apostrophe+$v+$Apostrophe }

	                $s+=$v+$ColSep
        	        $i++
                }
                $s=$s.Substring(0,$s.Length-$ColSep.Length)
                
                if ($csv -ne "") {
                    $s | Out-File $csv -Append -Encoding $Encoding
                } else { Write-Host $s }

            }

            $r.Close()
            $cnn.Close()

        }
        else { Write-Error $cnn.State }
    #}
    #catch { Write-Error $PSDebugContext }

}


function Get-OdbcData {
param(
[string]$CnnStr
,[string]$Query
,[string]$CSV=""   # CSV file
,[int]$Timeout=300
,[bool]$Quiet=$false
,[bool]$TextQualified=$true
,[string]$ColSep=","
,[string]$Encoding="ascii"
)

    $cnn = New-Object System.Data.Odbc.OdbcConnection($CnnStr)
    $cnn.Open()

    $Apostrophe=""
    if ($TextQualified) { $Apostrophe="""" }

    if ($cnn.State -ne "Closed") {

        $ts1=get-date 
        if ($CSV -ne "") { Write-Host $ts1.ToLongTimeString() "Connected. Executing query..." }

        $cmd = New-Object System.Data.Odbc.OdbcCommand($Query, $cnn)
        $cmd.CommandTimeout=$timeout

        try {
            $r = $cmd.ExecuteReader()
        }
        catch {
            $msg = $_.Exception.Message
            # if ($CSV -ne "") { Write-Host $msg  }
            throw $msg 
            return
        }

        $rc = 0

        # export data
        $ts2=get-date 
        $ts=(NEW-TIMESPAN –Start $ts1 –End $ts2).TotalSeconds
        if ($CSV -ne "") { Write-Host $ts2.ToLongTimeString() "Query completed in $ts seconds. Exporting data to $CSV" }

        # get field names
        $i=0
        $s=""
        while ($i -lt $r.FieldCount) {
            $n= $r.GetName($i)
            $s=$s+$Apostrophe+$n+$Apostrophe+$ColSep
            $i++
        }
        $s=$s.Substring(0,$s.Length-$ColSep.Length)
        if (!$Quiet) { Write-Host $s }

        if ($csv -ne "") {
            $s | Out-File $csv -Encoding $Encoding  # -Append 
        }

        # get rows
        while ($r.Read()) {   

            $rc++
            $i=0
            $s=""
            while ($i -lt $r.FieldCount) {
    
            $s=$s+$Apostrophe+$r[$i]+$Apostrophe+$ColSep
            $i++
            }
            $s=$s.Substring(0,$s.Length-$ColSep.Length)
            if (!$Quiet) { Write-Host $s }
            if ($csv -ne "") {
                $s | Out-File $csv -Append -Encoding $Encoding
            }
        }

        $ts2=get-date 
        $ts=(NEW-TIMESPAN –Start $ts1 –End $ts2).TotalSeconds
        if ($CSV -ne "") { Write-Host $ts2.ToLongTimeString() "Total runtime (seconds): $ts   Extracted rows: $rc " }

        $r.Close()
        $cnn.Close()

    }

}


function Export-CSVFile {
param(
     [string]$cnnStr=$(throw 'source connection string is required.')
    ,[string]$query=$(throw 'source sql query is required.')
    ,[string]$tgtPath=$(throw 'target path string is required.')
    ,[string]$tgtFileName=$(throw 'target file name is required.')
    )

    try {

        $csv = Join-Path $tgtPath $tgtFileName

        Write-AuditLog -LogState "S" -LogMessage $csv

        Get-OleDbData $cnnStr $query $csv 

        Write-AuditLog -LogState "F" -LogMessage $csv
        }
    catch {

        Write-AuditLog -LogState "E" -LogMessage $_.Exception.Message

    }

}


<#
function Encrypt-File {
param(
[string]$recipient 
,[string]$path
,[string]$destination = $path+".gpg"
,[string]$homedir = "E:/ETL/PGP_KEYS/HomeDir"
,[string]$workdir = "E:/ETL"
,[string]$gpgpath = "C:\Program Files (x86)\GnuPG\bin\gpg.exe"
)

Start-Process -Wait -NoNewWindow -FilePath $gpgpath -ArgumentList "-v","--always-trust","--homedir $homedir","--batch","--yes","-e","--recipient $recipient","--output $destination","$path" -WorkingDirectory $workdir

# $gpgpath  -v --always-trust --homedir $homedir --batch --yes -e --recipient $recipient --output $destination $path

# -WorkingDirectory $workdir

}
#>

function Get-SharepointList {
param(
  $listUri="" # e.g. /sites/company.sharepoint.com:/sites/SiteName/poas:/lists/unique-guid-list-identificator/" 
 ,$fields=""
 ,$token  # use https://www.jstoolset.com/jwt to troubleshoot
)


$query = "items"
if ($fields -ne "") {
    $query+="?expand=fields(select=$fields)"
}
# $query+="&top=$top"

$uri = "https://graph.microsoft.com/v1.0"+$listUri+$query
$headers = @{'Authorization' = $Token}
$rc = 0

while (![String]::IsNullOrEmpty($uri)) {

    $response = (Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers) 

    $uri=$response.'@odata.nextLink'
    $value=$response.value

    $results+=$value
    
    $rc+=$value.count
    Write-Progress -Activity "Got $rc records from $uri"

}

return $results

}

function Import-SharepointList {
param(
     [string]$listUri=$(throw 'sharepoint list uri is required.')
    ,[string]$fields=""
    ,[string]$token=$(throw 'bearer token is required.')
    ,[string]$tgtCnnStr=$(throw 'target connection string is required.')
    ,[string]$tgtTableName=$(throw 'target table name is required.')
    )

### TBD:
### 1. Need to persist/refresh access token somehow
### 3. Bulk insert would be faster
### 4. Add try/catch properly

    trap {
       Write-Warning "Error trapped"
       break
       }

    $items = Get-SharepointList -listUri $listUri -fields $fields -token $token

    Invoke-SQL $tgtCnnStr "TRUNCATE TABLE $tgtTableName"
    $rc=0
    $c=$items.Count
    foreach ($item in $items)
    {
      $Id=$item.id 
      $Created=$item.createdDateTime
      $LastModified=$item.lastModifiedDateTime
      $CreatedBy=$item.createdBy.user.email
      $LastModifiedBy=$item.lastModifiedBy.user.email

      #$Status=$item.fields.Status.Replace("'","''")
      #$CsgAccount=$item.fields.CSGAccount_x0023_.Replace("'","''")

      $sqlStr =  "INSERT INTO $tgtTableName "
      $sqlStrFields = "(Id, Created, CreatedBy, LastModified, LastModifiedBy"
      $sqlStrValues = " VALUES ('$Id','$Created','$CreatedBy','$LastModified','$LastModifiedBy'" 

      foreach($prop in $item.fields.PSObject.Properties) {
        if ($prop.Name -ne "@odata.etag") {
            $sqlStrFields += ",["+$prop.Name.Replace("'","''")+"]"
            $sqlStrValues += ",'"+$prop.Value.Replace("'","''")+"'"
        }
      }
      $sqlStr += $sqlStrFields + ") " + $sqlStrValues + ") "

      # Write-Host $sqlStr
      try {
        Invoke-SQL $tgtCnnStr $sqlStr
        }
      catch { $msg = $_
        Write-Warning "Error $msg in SQL: $sqlStr " 
        break 
      }

      $rc+=1
      $i=[int](100*$rc/$c)
      if ($rc % 20 -eq 0) {
        Write-Progress -Activity "Imported $rc of $c records into $tgtTableName" -Status "$i% Complete" -PercentComplete $i }
      # $item
      # $sqlStr
      # break
    }

# } catch {  Write-Host $Error[0].Exception.Message  }
}

 

function Export-SqlTableSchemaAndSampleDataToExcel {
param (
  $CnnStr
  , $schemaName = "dbo"
  , $tables
  , $FileName
  , $IncludeCopybook = $true
  , $TopCount = "50"
  , $IsXLVisible = $False
)

    $XL = New-Object -ComObject Excel.Application 
    $XL.Visible = $IsXLVisible
    $WB = $XL.Workbooks.Add()

    foreach ($tableName in $tables) {

        Write-Host "Exporting $SchemaName.$tableName"

        $sheetName = $tableName
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $WS = $WB.Worksheets.Add()
        $WS.Name = $sheetName 
        $row = 1
    
        if ($IncludeCopybook) {
            $records = Get-SQLData -sqlQuery "DECLARE @SqlQuery nvarchar(max) = '';
                select @SqlQuery = @SqlQuery+ ', ''' + DATA_TYPE	+ ISNULL('('+convert(varchar, CHARACTER_MAXIMUM_LENGTH) + ')', '') + ''' AS [' + COLUMN_NAME + ']' from INFORMATION_SCHEMA.COLUMNS where table_name = '$tableName' and table_schema = '$schemaName';
                set @SqlQuery = 'SELECT ' + stuff(@SqlQuery, 1, 1, '');
                EXEC SP_EXECUTESQL @SqlQuery;" -srcCnnStr $CnnStr

            $rc = 0
            foreach ($record in $records) {

                $rc += 1
                $column = 1
                foreach ($cell in $record) {
                    $WS.Cells.Item($row, $column) = $cell

                    if ($rc -eq 1) { 
                        $WS.Cells.Item($row, $column).Font.Bold = $True 
                        $WS.Cells.Columns($column).EntireColumn.Autofit() | Out-Null                
                        }

                    $column += 1
                }

                $row +=1
            }

            $row +=1
        }

        if ($TopCount -gt "0") {
            $records = Get-SQLData -sqlQuery "SELECT TOP $TopCount * FROM [$schemaName].[$tableName]"        
            $rc = 0
            foreach ($record in $records) {

                $rc += 1
                $column = 1
                foreach ($cell in $record) {
                    try {
                        $WS.Cells.Item($row, $column) = $cell }
                    catch { $WS.Cells.Item($row, $column) = "'"+$cell }

                    if ($rc -eq 1) {
                        $WS.Cells.Item($row, $column).Font.Bold = $True }
                    $column += 1

                }

                $row +=1
            }
        }

        $usedRange = $WS.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null
    }

    #Close Excel
    Write-Host "Saving $FileName"
    $WB.SaveAs($FileName)
    $WB.Close($true)
    $XL.Quit()
    Write-Host "Done"

}

Function Out-DataTable 
{
  $dt = new-object Data.datatable  
  $First = $true  
 
  foreach ($item in $input){  
    $DR = $DT.NewRow()  
    $Item.PsObject.get_properties() | foreach {  
      if ($first) {  
        $Col =  new-object Data.DataColumn  
        $Col.ColumnName = $_.Name.ToString()  
        $DT.Columns.Add($Col)       }  
      if ($_.value -eq $null) {  
        $DR.Item($_.Name) = $null #"[empty]"  
      }  
      elseif ($_.IsArray) {  
        $DR.Item($_.Name) =[string]::Join($_.value ,";")  
      }  
      else {  
        $DR.Item($_.Name) = $_.value  
      }  
    }  
    $DT.Rows.Add($DR)  
    $First = $false  
  } 
 
  return @(,($dt))
 
}

function Out-BulkSQL() {
param ($cnnString, $targetTable, $data, [bool] $truncate = $false, $timeout=300)

if ($truncate) {
  Invoke-SQL $cnnString "TRUNCATE TABLE $targetTable"
}

$bulkCopy = new-object ("Data.SqlClient.SqlBulkCopy") $cnnString
$bulkCopy.DestinationTableName = $targetTable 
$bulkCopy.BulkCopyTimeout = $timeout 
$bulkCopy.WriteToServer($data)

}


# example of AD accounts bulk extraction
function Import-AD() {
param(
$ADServerName = "dc.domain.com:3268"
, $filter = 'sAMAccountName -like "*"'
, $props = @('displayName','title','mail', 'department', 'AccountExpires', 'whenCreated')
, $cnnString = "Data Source=SQLDB; Integrated Security=True;Initial Catalog=ETL;"
, $targetTable = "[ADSI].[Accounts]"
)

$ad = (Get-ADUser -Filter $filter -Server $ADServerName -Properties $props) #  |  Select-Object objectGUID, comcastGUID, sAMAccountName, Enabled, AccountExpires, whenCreated, distinguishedName, displayName, title ....  )
$dt = ($ad | Out-DataTable )

Out-BulkSQL -cnnString $cnnString -targetTable $targetTable -data $dt -truncate $true -timeout 300

}

function Add-Quotes { param ([string]$Name) return "["+$Name+"]" }
function Remove-SqlInjectSymbols { param ([string]$Name) return $Name.replace(";", "").replace("[", "").replace("]", "").replace("'", "").replace(" UNION ", "") }

function Get-SqlTableCreateStmt {
param ([string]$CnnStr, [string]$DatabaseName, [string]$SchemaName, [string]$TableName
, [string]$CreateSchemaName = $SchemaName, [string]$CreateTableName = $TableName
, [bool]$ScriptIdentityColumns = $true
, [bool]$ScriptComputedColumns = $true
, [bool]$ScriptPK = $true
)

    # validate input
    $DatabaseName = Remove-SqlInjectSymbols($DatabaseName)
    $SchemaName = Remove-SqlInjectSymbols($SchemaName)
    $TableName = Remove-SqlInjectSymbols($TableName)

    # prepare create sql statement
    $QuotedDatabaseName = Add-Quotes($DatabaseName)

    $sql = "DECLARE @SqlQuery nvarchar(max) = '', @Constraint nvarchar(max) = '';
SELECT @SqlQuery = STRING_AGG(quotename(c.name) 
	+ " + (&{If($ScriptComputedColumns) {" CASE WHEN c.is_computed = 1 THEN ' AS ' + cc.definition ELSE"} ELSE {" CASE WHEN 1=0 THEN '' ELSE "}}) + "
        ' ' + t.name 
		+ CASE WHEN t.name in ('text','ntext','char','nchar','varchar','nvarchar') THEN ISNULL('('+case when c.max_length=-1 then 'max' else convert(varchar, c.max_length) end + ')', '')  ELSE '' END
		+ ISNULL(' DEFAULT ' + dc.definition, '')
		" + (&{If($ScriptIdentityColumns) {" + CASE WHEN c.is_identity = 1 THEN ' IDENTITY(' + CONVERT(VARCHAR, ic.seed_value) +',' + CONVERT(VARCHAR, ic.increment_value) +')' ELSE '' END "} Else {""}}) + "
		+ CASE WHEN c.is_nullable = 0 THEN ' NOT NULL' ELSE '' END
	END
	, ','+CHAR(10))	
	 WITHIN GROUP ( ORDER BY c.column_id ASC)   
FROM $QuotedDatabaseName.sys.columns c
INNER JOIN $QuotedDatabaseName.sys.types t ON t.system_type_id = c.system_type_id
LEFT JOIN $QuotedDatabaseName.sys.default_constraints dc ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
LEFT JOIN $QuotedDatabaseName.sys.computed_columns cc ON cc.object_id = c.object_id AND cc.column_id = c.column_id
LEFT JOIN $QuotedDatabaseName.sys.identity_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE object_name(c.object_id, db_id('$DatabaseName')) = '$TableName' 
and object_schema_name(c.object_id, db_id('$DatabaseName')) = '$SchemaName';
" + (&{If($ScriptPK) {"with pk as (select quotename(kc.name collate SQL_Latin1_General_CP1_CI_AS) + ' PRIMARY KEY ' + i.type_desc + ' (' as pk
, c.name, ic.index_column_id, ic.is_descending_key
FROM $QuotedDatabaseName.sys.key_constraints kc
INNER JOIN $QuotedDatabaseName.sys.indexes i ON i.object_id = kc.parent_object_id AND i.index_id = kc.unique_index_id
INNER JOIN $QuotedDatabaseName.sys.index_columns ic ON ic.object_id = kc.parent_object_id AND ic.index_id = kc.unique_index_id
INNER JOIN $QuotedDatabaseName.sys.columns c ON c.object_id = kc.parent_object_id AND c.column_id = ic.column_id
WHERE object_name(parent_object_id, db_id('$DatabaseName')) = '$TableName' and object_schema_name(parent_object_id, db_id('$DatabaseName')) = '$SchemaName'
), pk_c as (select STRING_AGG(quotename(name) + CASE WHEN is_descending_key=0 THEN ' ASC' ELSE ' DESC' END, ', ') WITHIN GROUP ( ORDER BY index_column_id ASC) as pk_c from pk)
select distinct @Constraint = pk.pk + pk_c.pk_c + ')' from pk, pk_c;"} ELSE {""}}) + "

set @SqlQuery = 'CREATE TABLE "+(Add-Quotes($CreateSchemaName))+"."+(Add-Quotes($CreateTableName))+" ('+ @SqlQuery
  + CASE WHEN ISNULL(@Constraint, '') != '' THEN ',' + CHAR(10) + 'CONSTRAINT ' + @Constraint + CHAR(10) ELSE '' END
  + ');';

SELECT @SqlQuery AS Stmt;"

    $output = Invoke-SQLQueryScalar -CnnStr $srcCnn -SqlStr $sql 

    return @(, ($output))
}

function Get-SqlTableFields {
param ([string]$CnnStr, [string]$DatabaseName, [string]$SchemaName, [string]$TableName
, [bool]$ExcludeIdentityColumns = $false
, [bool]$ExcludeComputedColumns = $false
)

    # validate input
    $DatabaseName = Remove-SqlInjectSymbols($DatabaseName)
    $SchemaName = Remove-SqlInjectSymbols($SchemaName)
    $TableName = Remove-SqlInjectSymbols($TableName)

    # prepare create sql statement
    $QuotedDatabaseName = Add-Quotes($DatabaseName)

    $sql = "DECLARE @SqlQuery nvarchar(max) = '';
SELECT @SqlQuery = STRING_AGG(quotename(c.name), ','+CHAR(10)) WITHIN GROUP (ORDER BY c.column_id ASC)
FROM $QuotedDatabaseName.sys.columns c
WHERE object_name(c.object_id, db_id('$DatabaseName')) = '$TableName' 
AND object_schema_name(c.object_id, db_id('$DatabaseName')) = '$SchemaName'
" + (&{If($ExcludeIdentityColumns) {" AND c.is_identity = 0"} ELSE {""} }) + (&{If($ExcludeComputedColumns) {" AND c.is_computed = 0"} ELSE {""} }) + ";

SELECT @SqlQuery AS Stmt;"

    # Write-Host $sql 
    $output = Invoke-SQLQueryScalar -CnnStr $srcCnn -SqlStr $sql 

    return @(, ($output))
}


function Copy-TableSchema {
param(
     [string]$srcCnnStr=$(throw 'source connection string is required.')
    ,[string]$srcDatabaseName=$(throw 'source sql query is required.')
    ,[string]$srcSchemaName=$(throw 'source sql query is required.')
    ,[string]$srcTableName=$(throw 'source sql query is required.')
    ,[string]$tgtCnnStr=$(throw 'target connection string is required.')
    ,[string]$tgtSchemaName=$(throw 'target table name is required.')
    ,[string]$tgtTableName=$(throw 'target table name is required.')
    ,[bool]$SchemaOnly=$true # means no data
    ,[bool]$ScriptIdentityColumns=$true
    ,[bool]$ScriptComputedColumns=$true
    ,[bool]$ScriptConstraints=$true
    )

    # (re)create table
    $sql = "DROP TABLE IF EXISTS $tgtSchemaName.$tgtTableName"
    Write-Host $sql
    Invoke-SQL -CnnStr $tgtCnnStr -SqlStr $sql 

    $createSql = Get-SqlTableCreateStmt -CnnStr $srcCnnStr -DatabaseName $srcDatabaseName -SchemaName $srcSchemaName -TableName $srcTableName `
        -ScriptIdentityColumns $ScriptIdentityColumns -ScriptComputedColumns $ScriptComputedColumns -ScriptPK $ScriptConstraints `
        -CreateSchemaName $tgtSchemaName -CreateTableName $tgtTableName 
    
    Write-Host $createSQL
    Invoke-SQL -CnnStr $tgtCnnStr -SqlStr $createSql

    if (!$SchemaOnly) {
        # copy data
        $srcSQL = "SELECT * FROM $srcDatabaseName.$srcSchemaName.$srcTableName"
        Write-Host "$srcSQL INTO $tgtSchemaName.$tgtTableName"
        Import-SqlData -srcCnnStr $srcCnn -tgtCnnStr $AzureCnn -srcSql $srcSQL -tgtTableName ($tgtSchemaName+"."+$tgtTableName)
    }
}

function Get-FixIdentityColumnScript {
param([string]$srcCnnStr = $(throw 'is required'), [string]$srcDatabaseName = $(throw 'is required'), [string]$srcSchemaName = $(throw 'is required'), [string]$srcTableName = $(throw 'is required')
,[string]$tgtCnnStr = $(throw 'is required'), [string]$tgtSchemaName = $srcSchemaName, [string]$tgtTableName = $srcTableName
)

[string] $sql1 = Get-SqlTableCreateStmt -CnnStr $srcCnnStr -DatabaseName $srcDatabaseName  -SchemaName $srcSchemaName -TableName $srcTableName `
-CreateSchemaName $tgtSchemaName -CreateTableName "TMP_$tgtTableName" -ScriptIdentityColumns $true -ScriptComputedColumns $true -ScriptPK $true 


[string] $fields = Get-SqlTableFields -CnnStr $srcCnnStr -DatabaseName $srcDatabaseName  -SchemaName $srcSchemaName -TableName $srcTableName -ExcludeComputedColumns $true
[string] $sql2 = "IF EXISTS (SELECT NULL FROM [$tgtSchemaName].[$tgtTableName])
INSERT INTO [$tgtSchemaName].[TMP_$tgtTableName] (" + $fields + "
)
SELECT " + $fields + "
FROM [$tgtSchemaName].[$tgtTableName] TABLOCKX;
"

[string] $sql = $sql1
if ($sql1.Contains(" IDENTITY(")) {
    $sql += "`n" + "SET IDENTITY_INSERT [$tgtSchemaName].[TMP_$tgtTableName] ON;"
    $sql2 += "`n" + "SET IDENTITY_INSERT [$tgtSchemaName].[TMP_$tgtTableName] OFF;"
} 

return $sql + "`n`n" + $sql2 + "`n"


}


function Copy-AzSqlTableData {
param([string] $SrcServer
, [string] $TgtServer
, [string] $SrcDatabaseName
, [string] $TgtDatabaseName = $SrcDatabaseName
, [string] $SrcTableName
, [string] $TgtTableName = $SrcTableName
, [string] $Query = "SELECT * FROM $SrcTableName;"
, [bool] $Identity = $false
)

$Token = (Get-AzAccessToken -ResourceUrl https://database.windows.net).Token
$Table = Invoke-Sqlcmd -ServerInstance $SrcServer -Database $SrcDatabaseName -AccessToken $Token -Query $Query -MaxCharLength 8000;

$SrcCount = $Table.Count 
If ($SrcCount -gt 0) {

    Write-Host "Copying $SrcDatabaseName.$SrcTableName ($SrcCount rows) from $SrcServer to $TgtDatabaseName.$TgtTableName on $TgtServer"

    Invoke-Sqlcmd -ServerInstance $TgtServer -Database $TgtDatabaseName -AccessToken $Token -Query "DELETE FROM $TgtTableName;"

    $count = 0
    foreach ($row in $Table) {
       
        $cols = ""
        $vals = ""
        foreach ($col in $row.Table.Columns) {
            $cols += $col.ColumnName + ", "
            $vals += "'" + ([string] $row[$col.ColumnName]).Replace("'", "''").Replace('$',"' + CHAR(36) + '") + "', "
        }

        $cols = $cols.Substring(0, $cols.Length - 2)
        $vals = $vals.Substring(0, $vals.Length - 2)
        $sql = "INSERT INTO $TgtTableName ($cols) VALUES ($vals)"
        if ($Identity) { $sql = "SET IDENTITY_INSERT $TgtTableName ON; " + $sql }
        try {
            $result = Invoke-Sqlcmd -ServerInstance $TgtServer -Database $TgtDatabaseName -AccessToken $Token -Query $sql
            }
        catch {
            Write-Error $sql
            Write-Error $_
            }
        $count += 1
    }

    Write-Host "$count rows copied"

} else { Write-Host "The source $SrcTableName table at $SrcServer ($SrcDatabaseName db) is empty" }

}

Export-ModuleMember *-*
