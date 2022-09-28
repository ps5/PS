<#
.SYNOPSIS
  Various T-SQL helper scripts

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Paul Shiryaev <ps@paulshiryaev.com>
  
#>

function Invoke-SQL {
    param(
       [string]$CnnStr=$(throw 'connection string is required.'),
       [string]$SqlStr=$(throw 'command string is required.'),
       [int]$Timeout=60 # the wait time (in seconds) before terminating the attempt to execute a command and generating an error.
       )

    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = $CnnStr   
    $cnn.Open()

    $cmd = $cnn.CreateCommand()
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $SqlStr
    $cmd.CommandTimeout = $Timeout
    $rc = $cmd.ExecuteNonQuery();

    $cnn.Close()

}


function Invoke-SQLQueryScalar {
    param(
       [string]$CnnStr=$(throw 'connection string is required.'),
       [string]$SqlStr=$(throw 'command string is required.'),
       [int]$CommandTimeout = 120
       )

    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = $CnnStr   
    $cnn.Open()

    $cmd = $cnn.CreateCommand()
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $SqlStr
    $cmd.CommandTimeout = $CommandTimeout
    $result = $cmd.ExecuteScalar()

    $cnn.Close()
    return $result

}


# Updates stored domain account credentials and linked server logins
# 

Function Update-Credentials {
param (
 [string]$ServerName,
 [string]$AccountName,
 [string]$Password,
 $testrun=$true)


    Write-Host "Checking $ServerName"

    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=TSQL.PS"
    $cnn.Open()

    $sql = "SELECT a.name as ls FROM sys.Servers a LEFT JOIN sys.linked_logins b ON b.server_id = a.server_id where remote_name = '$AccountName' and is_linked = 1" 
    $sql = $sql + " and b.modify_date < getdate()-0.05" ## exclude last hour

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $command = $cnn.CreateCommand()
    $command.CommandText = $sql
    $adapter.SelectCommand = $command
    $dataset = New-Object System.Data.DataSet
    $rc = $adapter.Fill($dataset);
    if ($rc -gt 0) {

        Write-Host "Found $rc linked servers to update..."

        foreach($row in $dataset.Tables[0].Rows) {        
            $ls = $row["ls"]
            Write-Host "* Updating $ls"

            $cmd = $cnn.CreateCommand()
            $cmd.CommandType = [System.Data.CommandType]::StoredProcedure
            $cmd.CommandText = "master.dbo.sp_addlinkedsrvlogin"
            $cmd.Parameters.AddWithValue("@rmtsrvname", $ls) | Out-Null
            $cmd.Parameters.AddWithValue("@rmtuser", $AccountName) | Out-Null
            $cmd.Parameters.AddWithValue("@rmtpassword", $Password) | Out-Null
            $cmd.Parameters.AddWithValue("@useself", 'False') | Out-Null
            $cmd.Parameters.AddWithValue("@locallogin", $null) | Out-Null

            if (!($testrun)) { $rc = $cmd.ExecuteNonQuery(); }


        }

    }



    $sql = "select name as cred from sys.credentials where credential_identity = '$AccountName'"
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $command = $cnn.CreateCommand()
    $command.CommandText = $sql
    $adapter.SelectCommand = $command
    $dataset = New-Object System.Data.DataSet
    $rc = $adapter.Fill($dataset);
    if ($rc -gt 0) {

        Write-Host "Found $rc credentials to update..."

        foreach($row in $dataset.Tables[0].Rows) {        
            $cred = $row["cred"]
            Write-Host "* Updating $cred"

            $cmd = $cnn.CreateCommand()
            $cmd.CommandType = [System.Data.CommandType]::Text
            $cmd.CommandText = "ALTER CREDENTIAL [$cred] WITH IDENTITY = '$AccountName', SECRET = '$Password'"
            if (!($testrun)) { $rc = $cmd.ExecuteNonQuery() }

        }

    }


    $cnn.Close()
    
 }



# SAMPLE RUN
#
# $login = 'Domain\ServiceAccountName'
# $password = 'newsecretpassword'
# Update-CredentialsAllServers -config "servers.json" -login $login -password $password -testrun $True
# 
function Update-CredentialsAllServers {
param ([string]$config=$(throw 'servers.json configuration file is required')
,[string]$login
,[string]$password
,$testrun=$true)

    try {

        $json = (Get-Content $config | ConvertFrom-Json)

	Write-Host "Updating credentials $login"
	Write-Host $password
	if (!($testrun)) { Write-Host "*** LIVE RUN ***" }

        foreach ($server in $json.servers) {

    		if ("sql" -in $server.roles) {
                Update-Credentials $server.fqdn $login $password $testrun
            }
		}


    }
    catch
    {
        Write-Host $_.Exception.Message
    }
}


# SQL SERVER AGENT RELATED LOGGING
function Import-SqlAgentLogs {
param(
    [string]$ServerName=$(throw 'server name is required.')
    ,$tgtCnnStr=$TargetCnnStr    
    )

    $srcCnnStr = "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=LOGGING.PS"

    ### get last log record id
    $LastID = Invoke-SQLQueryScalar -CnnStr $tgtCnnStr -SqlStr "select isnull(max(instance_id), 0) from ADMIN.logs.sqlagent (NOLOCK) where server_name = '$ServerName'"
    Write-Host "$ServerName - last id: $LastID"
    ## import jobs history
    $sql = Get-SqlAgentLogsQuery($LastID)    
    Import-SqlData -srcCnnStr $srcCnnStr -srcSql $sql -tgtCnnStr $tgtCnnStr -tgtTableName "logs.sqlagent"

    ## import running jobs
    Invoke-SQL -CnnStr $tgtCnnStr -SqlStr "delete from logs.sqlagent_running where server_name = '$ServerName'"
    $sql = Get-SqlAgentJobsRunningQuery
    Import-SqlData -srcCnnStr $srcCnnStr -srcSql $sql -tgtCnnStr $tgtCnnStr -tgtTableName "logs.sqlagent_running"

    ## import jobs
    Invoke-SQL -CnnStr $tgtCnnStr -SqlStr "delete from logs.sqlagent_jobs where server_name = '$ServerName'"
    $sql = Get-SqlAgentJobsQuery
    Import-SqlData -srcCnnStr $srcCnnStr -srcSql $sql -tgtCnnStr $tgtCnnStr -tgtTableName "logs.sqlagent_jobs"

}

function Get-SqlAgentLogsQuery($LastID) {

    return "select jh.server as server_name
, jh.instance_id
, j.category_id
, jh.job_id
, jh.step_id
, jh.run_status
, jh.run_date
, jh.run_time
, jh.run_duration
, jh.sql_message_id
, jh.sql_severity
, jh.retries_attempted

, c.name as category_name
, j.name as job_name
, jh.step_name
, jh.message

from msdb.dbo.sysjobhistory jh (nolock) 
inner join msdb.dbo.sysjobs j (nolock) on j.job_id = jh.job_id
left join msdb.dbo.sysjobsteps js (nolock) on js.job_id = jh.job_id and js.step_id = jh.step_id
left join msdb.dbo.syscategories c (nolock) on c.category_id = j.category_id
where jh.instance_id > " + $LastID + "  
and not (j.name like 'DBA%' and jh.run_duration < 30)
order by jh.instance_id asc"

}


function Get-SqlAgentJobsQuery {

    return "SELECT @@SERVERNAME as server_name
		, j.job_id
		, j.name as job_name
		, j.description as job_description
		, c.name as category_name
		, js.count_of_steps
		, left(so.name, 50) as job_operator
		, so.netsend_address as operator_netsend_address
		, so.email_address as operator_email_address
		, j.enabled as job_enabled
		, j.date_created as job_created
		, j.date_modified as job_modified
		, j.version_number as job_version_number
 		FROM msdb.dbo.sysjobs j (nolock) 
		LEFT JOIN msdb.dbo.syscategories c (nolock) on c.category_id = j.category_id
		LEFT JOIN msdb.dbo.sysoperators (nolock) so on so.id=j.notify_email_operator_id 
		OUTER APPLY (SELECT count(*) as count_of_steps FROM msdb.dbo.sysjobsteps js (nolock) WHERE j.job_id = js.job_id) js
		;"

}



function Get-SqlAgentJobsRunningQuery {

    return "SELECT @@SERVERNAME as server_name		
    ,jc.*
	,js.step_name 
	,js.subsystem
	,js.command
FROM (
	SELECT
		c.name as category_name
		,ja.job_id
		,j.name AS job_name
		,ja.start_execution_date as job_start_time
		,ja.last_executed_step_id
		,jh.run_status as last_executed_step_status
		, current_step_id = case when ja.last_executed_step_id is null then j.start_step_id else
			case jh.run_status 
			when 0 /* failed */ then case js.on_fail_action when 1 then 0 when 2 then 0 when 3 then js.step_id+1 when 4 then js.on_fail_step_id else 0 end
			when 1 /* succeeded */ then case js.on_success_action when 1 then 0 when 2 then 0 when 3 then js.step_id+1 when 4 then js.on_success_step_id else 0 end
			when 2 /* retry */ then js.step_id
			when 3 /* cancelled */ then 0
			when 4 /* in progress */ then js.step_id
			else 0 end
		end
		,isnull(dateadd(second, jh.run_time % 100, dateadd(minute, jh.run_time / 100 % 100, dateadd(hour, jh.run_time / 10000, convert(datetime, convert(varchar, jh.run_date))))), ja.start_execution_date) AS step_start_time
		/* ,ja.next_scheduled_run_date */
	FROM msdb.dbo.sysjobactivity ja (nolock) 
	INNER JOIN msdb.dbo.sysjobs j (nolock) ON ja.job_id = j.job_id
	LEFT JOIN msdb.dbo.syscategories c (nolock) on c.category_id = j.category_id
	LEFT JOIN msdb.dbo.sysjobsteps js (nolock) ON ja.job_id = js.job_id AND ja.last_executed_step_id = js.step_id
	outer apply (select top 1 * from msdb.dbo.sysjobhistory jh (nolock) where ja.job_id = jh.job_id and ja.last_executed_step_id = jh.step_id order by instance_id desc) jh
	WHERE ja.session_id = (SELECT TOP 1 session_id FROM msdb.dbo.syssessions (nolock) ORDER BY agent_start_date DESC)
	AND start_execution_date is not null
	AND stop_execution_date is null
) jc
INNER JOIN msdb.dbo.sysjobsteps js (nolock) ON jc.job_id = js.job_id AND jc.current_step_id = js.step_id
order by job_name;"

}

function Resume-FailedSqlAgentJobs {
    param(
       [string]$ServerName=$(throw 'server name is required.')
       )

    $criticalerror = 0
    # $verbose = 0
       
    try {
        # if ($verbose) { Write-Host $(Get-Date) "$ServerName" }

        $sql = Get-QueryResumeAbortedSqlAgentJobsDueToRestart 
        Invoke-SQL -CnnStr "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=SQLAGENT.PS" `
            -SqlStr $sql 


        $sql = Get-QueryStartMissedSqlAgentJobsDueToRestart 
        Invoke-SQL -CnnStr "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=SQLAGENT.PS" `
            -SqlStr $sql 
            
        
    }

    catch {
        $criticalerror = 1
        $ErrorMsg = $ServerName+": "+$_.Exception.Message # $Error[0].ToString()
        Write-Error $ErrorMsg
        Write-AuditLog -LogState 'E' -LogMessage $ErrorMsg
        break
    }

}

function Get-QueryResumeAbortedSqlAgentJobsDueToRestart() {

    $ServerName = $env:COMPUTERNAME
    $ScriptPath = Join-Path $PSScriptRoot "TSQL.psm1"


    return "/* 
    -- working version 2/16/2021
    -- restarts jobs that got aborted in running by a server restart and got stuck in running mode
    -- required permissions:

use master
grant view server state to [domain\localmachine$]
use msdb
alter role SQLAgentOperatorRole add member [domain\localmachine$]

    */
    
    set nocount on
	
declare @starttime datetime = (SELECT sqlserver_start_time - (getutcdate() - getdate()) FROM sys.dm_os_sys_info); /* get uptime */
if @starttime > getdate() set @starttime = @starttime + (getutcdate() - getdate()); /* bugfix for HP Proliant UTC bug */
	
if datediff(minute, @starttime, getdate()) < 120 /* did server restart within the last 2 hours */
begin

	declare @cmds table (cmd nvarchar(max), note nvarchar(max));
	insert into @cmds
	select distinct cmd = 'exec msdb.dbo.sp_start_job @job_id = ''' + convert(varchar(max), x.job_id) + ''', @step_name = ''' + s.step_name + ''''
	, note = 'Restarted aborted ' + x.job_name + ' (step ' + convert(varchar, x.run_step_id) + ': ' + s.step_name + ')'
	from (
	select run_step_id = case when jh.run_status = 1 then /* run next step */
		case when on_success_step_id = 0 and on_success_action = 3 then jh.step_id+1 else on_success_step_id end 
		else jh.step_id /* retry last step */ end
	, jh.job_id
	, j.name as job_name
	from [msdb].[dbo].sysjobhistory jh
	inner join [msdb].[dbo].sysjobsteps js on js.job_id=jh.job_id and js.step_id=jh.step_id
	inner join msdb.dbo.sysjobs j on j.job_id=jh.job_id 
	where jh.instance_id in (select last_instance_id from ( SELECT job_id, max(instance_id) last_instance_id FROM [msdb].[dbo].sysjobhistory (nolock) group by job_id) x) 
	and jh.step_id <> 0
	and not exists (select null FROM msdb.dbo.sysjobactivity ja (nolock) WHERE ja.job_id = jh.job_id
	and ja.session_id = (SELECT max(session_id) FROM msdb.dbo.syssessions (nolock) WHERE agent_start_date >= GETDATE()-365)
	AND start_execution_date is not null AND stop_execution_date is null
	) /* is not running */
	and j.enabled=1 /* job enabled */
	) x
	inner join [msdb].[dbo].sysjobsteps s on s.job_id=x.job_id and s.step_id=x.run_step_id

	if @@rowcount > 0
	begin
		declare @cmd nvarchar(max) = (select top 1 cmd from @cmds order by cmd);
		while @cmd <> ''
		begin
			declare @msg nvarchar(max) = (select top 1 note from @cmds where cmd = @cmd);
			print @msg;
			exec (@cmd);

			exec admin.logs.sp_Audit_Proc '[$ServerName] $ScriptPath', 'W', @msg;
			set @cmd = (select top 1 cmd from @cmds where cmd > @cmd order by cmd);
		end
	end

end	


"

}

function Get-QueryStartMissedSqlAgentJobsDueToRestart() {
param($JobFilterName = 'Daily')

    $ServerName = $env:COMPUTERNAME
    $ScriptPath = Join-Path $PSScriptRoot "TSQL.psm1"

    return "/* 
    -- working version 6/24/2022
    -- restarts daily jobs that didn't run because of the reboot (DELL firmware UTC bug)
    */
        
    set nocount on
	
	declare @starttime datetime = (SELECT sqlserver_start_time - (getutcdate() - getdate()) FROM sys.dm_os_sys_info); /* get uptime */
	if @starttime > getdate() set @starttime = @starttime + (getutcdate() - getdate()); /* bugfix for HP Proliant UTC bug */
	
	if datediff(minute, @starttime, getdate()) < datediff(hour, GETDATE(), GETUTCDATE()) * 60 /* did server restart within the last 4-5 hours */
	begin

    if object_id('tempdb..#jobs') is not null drop table #jobs
    
    select job_id, sj.name
    , sc.category_name 
    , is_daily = case when (sj.name like '%Daily%' /* decode start time from the name */
        and substring(sj.name, patindex('%DAILY%', sj.name) + 6, 4) < substring(convert(varchar, getdate(), 108), 1, 2) + substring(convert(varchar, getdate(), 108), 4, 2)
        ) then 1 else -- 0 
            case when sch.name is not null then 1 else 0 end /* missed the schedule */
        end
    into #jobs 
    from msdb.dbo.sysjobs sj
    outer apply (select sc.name as category_name from msdb.dbo.syscategories sc where sc.category_id = sj.category_id) sc
    outer apply (select top 1 ss.name, next_run_date from msdb.dbo.sysjobschedules sjs 
    inner join msdb.dbo.sysschedules ss on ss.schedule_id = sjs.schedule_id
    where sj.job_id = sjs.job_id and ss.enabled = 1 and convert(varchar, getdate(), 112) between ss.active_start_date	and ss.active_end_date
    and ss.freq_type = 4 and ss.freq_interval = 1 
    and next_run_date > convert(varchar, getdate(), 112) /* job missed the schedule */
    ) sch
    where sj.enabled = 1 and sj.name not like 'DEL%' and ( sj.name like '%"+$JobFilterName+"%' )
    order by category_name, name
    
     
     DECLARE @xp_results TABLE (
        job_id UNIQUEIDENTIFIER NOT NULL
        ,running INT NOT NULL)

	INSERT INTO @xp_results SELECT
		ja.job_id
		,Running=1
		/*,j.name AS job_name
		,ja.start_execution_date as job_start_time
		,ja.last_executed_step_id*/
	FROM msdb.dbo.sysjobactivity ja (nolock) 
	INNER JOIN msdb.dbo.sysjobs j (nolock) ON ja.job_id = j.job_id
	-- LEFT JOIN msdb.dbo.sysjobsteps js (nolock) ON ja.job_id = js.job_id AND ja.last_executed_step_id = js.step_id
	--outer apply (select top 1 * from msdb.dbo.sysjobhistory jh (nolock) where ja.job_id = jh.job_id and ja.last_executed_step_id = jh.step_id order by instance_id desc) jh
	WHERE ja.session_id = (SELECT max(session_id) FROM msdb.dbo.syssessions (nolock) WHERE agent_start_date >= GETDATE()-365)
	AND start_execution_date is not null
	AND stop_execution_date is null;
    
    
    declare @job_id UNIQUEIDENTIFIER, @run_status int,@step_id INT, @name VARCHAR(max), @step_name VARCHAR(max), @is_running bit, @is_daily bit;
    declare @instance_id int, @previous_step_id INT;
    declare @category_name varchar(255), @previous_category_name varchar(255) = '', @previous_step_name varchar(255) = '';
    declare @restart_required bit;
    
    set @job_id = (select top 1 job_id from #jobs order by name);
    
    while @job_id is not null
    begin
        set @restart_required = 0;
        set @name = (select top 1 name from #jobs where job_id = @job_id);
        set @category_name = (select top 1 category_name from #jobs where name = @name);
        set @is_daily = (select top 1 is_daily from #jobs where name = @name);
        declare @msg varchar(255) = 'Restarted ' + @name
    
        SELECT TOP 1 @run_status = run_status, @step_name = step_name, @step_id = step_id, @instance_id = instance_id 
        FROM [msdb].[dbo].sysjobhistory (nolock)
        WHERE run_date = convert(INT, convert(VARCHAR, getdate(), 112))
            AND job_id = @job_id
        ORDER BY instance_id DESC;
    
        SELECT TOP 1 @previous_step_id = step_id, @previous_step_name = step_name 
        FROM [msdb].[dbo].sysjobhistory (nolock)
        WHERE run_date = convert(INT, convert(VARCHAR, getdate(), 112)) AND job_id = @job_id and instance_id < @instance_id 
        ORDER BY instance_id DESC;
    
    
        set @is_running = isnull((select Running from @xp_results where job_id = @job_id), 0);
    
        /* print '-- ' + @name + ': ' + isnull(convert(varchar, @run_status ), 'N/A') + ' - ' + isnull(@step_name, 'N/A')  */
    
        declare @sql varchar(max) = 'exec msdb.dbo.sp_start_job ''' + @name + '''';
    
        /* for daily jobs only: */
        if @is_daily = 1
            begin
                if @run_status is null and isnull(@is_running, 0) <> 1 /* did not start */
                begin
                    set @restart_required = 1;
                    set @sql = @sql + ' -- DID NOT RUN'
                    set @msg = @msg + ' -- DID NOT RUN'
                end
            end
    
        /* for all jobs: */
		/*
        if @run_status in (2,3)  and @is_running = 0 /* retry, stopped */
            begin
                if @step_name = '(Job outcome)' set @step_name = @previous_step_name 
                set @restart_required = 1;
                set @sql = @sql + case when @step_name <> '(Job outcome)' then ', @step_name = ''' + @step_name + '''' else '' end; 
                set @sql = @sql + ' -- STOPPED'
                set @msg = @msg + ' -- STOPPED'		
            end
        else if (@run_status = 1 and  isnull(@step_name, 'N/A') <> '(Job outcome)') and @is_running = 0 /* aborted halfway by a restart */
            begin
                set @restart_required = 1;
                set @step_id = (select case when on_success_step_id = 0 and on_success_action = 3 then step_id+1 else on_success_step_id end 
                    from msdb.dbo.sysjobsteps where job_id = @job_id and step_id = @step_id);
                set @step_name = (select step_name from  msdb.dbo.sysjobsteps where job_id = @job_id and step_id = @step_id);
                set @sql = @sql + ', @step_name = ''' + @step_name + '''';
                set @sql = @sql + ' -- ABORTED'
                set @msg = @msg + ' -- ABORTED'
            end
    
	*/

        if @restart_required = 1 and @sql is not null
        begin
            if @previous_category_name != @category_name
                begin
                    set @previous_category_name = @category_name
                    print '/* ' + @category_name + ' */'
                end
            print @sql
            --exec (@sql)
            
			exec admin.logs.sp_Audit_Proc '[$ServerName] $ScriptPath', 'W', @msg;
    
        end
    
        set @job_id = (select top 1 job_id from #jobs where name > @name order by name);
        set @run_status = null;
    
    end

	end
    
    "
}



function Import-SqlData {
param(
     [string]$srcCnnStr=$(throw 'source connection string is required.')
    ,[string]$srcSql=$(throw 'source sql query is required.')
    ,[string]$tgtCnnStr=$(throw 'target connection string is required.')
    ,[string]$tgtTableName=$(throw 'target table name is required.')
    )

    $tgtCnn = New-Object System.Data.SqlClient.SqlConnection($tgtCnnStr)
    $tgtCnn.Open()

    $srcCnn = New-Object System.Data.SqlClient.SqlConnection($srcCnnStr)
    $srcCnn.Open()

    $cmd = $srcCnn.CreateCommand()
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $srcSql
    $dr = $cmd.ExecuteReader()
    # Write-Host $sql

    [int]$rc=0
    while ($dr.Read()) {

        $sqlh = "insert into $TgtTableName ("
        $sqlv = " values ("

        for ($i=0; $i -lt $dr.FieldCount; $i++) {

            $sqlh += $dr.GetName($i) + ","            
            $vstr = $dr.GetValue($i).ToString().Replace("'","''")
            $sqlv += "'"+$vstr+"',"            
        }

        $sqlh = $sqlh.Substring(0, $sqlh.Length-1) + ") "
        $sqlv = $sqlv.Substring(0, $sqlv.Length-1) + ") "

        $cmd = $tgtCnn.CreateCommand()
        $cmd.CommandType = [System.Data.CommandType]::Text
        $sql = $sqlh + $sqlv
        $cmd.CommandText = $sql
        # Write-Host $sql
        $rc += $cmd.ExecuteNonQuery()

    }
    Write-Host "Imported $rc records into $tgtTableName"

    $srcCnn.Close()
    $tgtCnn.Close()

}

function Get-SQLData {
param ($sqlQuery=$(throw 'sql query is required')
, $srcCnnStr =$(throw 'sql connection string is required')
, $SkipHeader=$false)

    $srcCnn = New-Object System.Data.SqlClient.SqlConnection($srcCnnStr)
    $srcCnn.Open()

    $cmd = $srcCnn.CreateCommand()
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $sqlQuery
    $dr = $cmd.ExecuteReader()

    $results = @()

    if (!$SkipHeader) {

        $row = @()
        for ($i=0; $i -lt $dr.FieldCount; $i++) {
            $row  += $dr.GetName($i)
        }
        $results +=  ,($row )

    }


    while ($dr.Read()) {

        $row = @()
        for ($i=0; $i -lt $dr.FieldCount; $i++) {
            $row  += $dr.GetValue($i)
        }
        $results +=  ,($row )

    }

    Write-Output (,$results)

    $srcCnn.Close()

}

function ConvertTo-Sanitized([string]$str) {
    return $str.Replace(';', '').Replace('[', '').Replace(']', '').Replace('*', '').Replace('?', '').Replace('(', '').Replace(')', '').Replace('/', '').Replace('\', '').Replace("'", '')
    }

function Find-Where {
    param(
       [string]$ServerName=$(throw 'sql server name is required.'),
       [string]$PatternStr1=$(throw 'pattern string to search is required.'),
       [string]$PatternStr2="", ## optional
       [boolean]$IncludeCode=$false
       )


       $PatternStr1 = ConvertTo-Sanitized $PatternStr1
       $PatternStr2 = ConvertTo-Sanitized $PatternStr2
       $results = @()

       Write-Host "Searching for '$PatternStr1' '$PatternStr2' on $ServerName"

       $cnnStr = "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=TSQL.PS"
       $sqlQuery = "select name as dbname, database_id from master.sys.databases where state_desc = 'ONLINE' and name not in ('msdb','tempdb') order by name"
       $databases = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
       foreach ($db in $databases) {
        
            # Write-Host $db[0]

            # search in code
            $sqlQuery = 'SELECT DISTINCT ''' + $db[0] + '.'' + object_schema_name(sc.id, ' + $db[1] + ') + ''.'' + obj.Name /*, convert(nvarchar(max),sc.TEXT)*/ ' `
                + ' FROM ' + $db[0] + '.sys.syscomments (nolock) sc INNER JOIN ' + $db[0] + '.sys.objects (nolock) obj ON sc.Id = obj.OBJECT_ID ' `
                + ' WHERE TYPE IN (''P'',''IF'',''FN'',''TF'', ''V'', ''SN'') ' `
                + ' AND (replace(replace(sc.TEXT, ''['', ''''), '']'','''') LIKE ''%' + $PatternStr1 + '%'' AND sc.TEXT LIKE ''%' + $PatternStr2 + '%'')'

            $result = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
            if ($result) {
                $results += $result
                }

            # search in synonyms
            $sqlQuery = 'SELECT DISTINCT ''SYNONYM: '' + object_schema_name(object_id, ' + $db[1] + ') + ''.'' + name, ' `
              + ''' for '' + base_object_name from ' + $db[0] + '.sys.synonyms (nolock) ' `
              + ' where (base_object_name LIKE ''%' + $PatternStr1 + '%'' AND base_object_name LIKE ''%' + $PatternStr1 + '%'')' `
              + ' or (name LIKE ''%' + $PatternStr1 + '%'' AND name LIKE ''%' + $PatternStr1 + '%'')'

            $result = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
            if ($result) {
                $results += $result
                }
            
       }

       # check msdb.dbo.sysjobs
        $sqlQuery = 'select j.name + '' - '' + convert(varchar, s.step_id) + ''. '' + s.step_name, s.subsystem + '': '' + s.command ' `
            + 'from msdb.dbo.sysjobs j (nolock) inner join msdb.dbo.sysjobsteps s (nolock) on s.job_id = j.job_id ' `
            + ' where ( replace(replace(s.command , ''['', ''''), '']'','''') LIKE ''%' + $PatternStr1 + '%''' `
	        + ' and replace(replace(s.command , ''['', ''''), '']'','''') LIKE ''%' + $PatternStr2 + '%'' )'

        $result = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
        if ($result) {
            $results += $result
            }
             
       return $results
       
}

function Find-Column {
    param(
       [string]$ServerName=$(throw 'sql server name is required.'),
       [string]$ColumnName=$(throw 'column name string is required.')
       )

       $ColumnName = ConvertTo-Sanitized $ColumnName
       $results = @()

       Write-Host "Searching for '$ColumnName' on $ServerName"

       $cnnStr = "Persist Security Info=False;Integrated Security=true;Initial Catalog=master;server=$ServerName;Application Name=TSQL.PS"
       $sqlQuery = "select name as dbname, database_id from master.sys.databases where state_desc = 'ONLINE' and name not in ('msdb','tempdb') order by name"
       $databases = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
       foreach ($db in $databases) {
        
            # Write-Host $db[0]

            # search in schema
            $sqlQuery = 'SELECT DISTINCT ''' + $db[0] + '.'' + object_schema_name(sc.id, ' + $db[1] + ') + ''.'' + obj.Name
 + ''   '' + convert(Varchar, sc.colorder) + '': ['' + sc.name + ''] as '' + t.name 
  + isnull('' '' + case when sc.collation is not null then ''('' + convert(varchar, sc.length) + '') COLLATE '' + sc.collation collate SQL_Latin1_General_CP1_CI_AS else null end, '''') 
  + '' '' + case when sc.isnullable = 0 then ''NOT '' else '''' end + ''NULL''
  FROM ' + $db[0] + '.sys.syscolumns sc INNER JOIN ' + $db[0] + '.sys.objects obj ON sc.Id = obj.OBJECT_ID
  INNER JOIN ' + $db[0] + '.sys.systypes t ON t.xtype = sc.xtype
  WHERE sc.name = ''' + $ColumnName + ''' ';

            #write-host $sqlQuery
            #return
  
            $result = Get-SQLData -sqlQuery $sqlQuery -srcCnnStr $cnnStr -SkipHeader $true 
            if ($result) {
                $results += $result
                }

            
       }

       return $results
       
}

Export-ModuleMember *-*
