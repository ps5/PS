<#
.SYNOPSIS
  Export and commit to GIT: schema, SSAS cubes, SSIS packages, SQL Agent jobs

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Created by Paul Shiryaev <ps@shiryaev.net> 
  github.com/ps5
#>

param(
    # path to .ssh parent folder (git home path)
    [parameter(Position=0,Mandatory=$false)][string]$HomePath = "E:\GIT"
    , [parameter(Position=1,Mandatory=$false)][string]$DomainName = @("CABLE\","@comcast.com")
    , [parameter(Position=2,Mandatory=$false)][string]$GitAuthor = "Wall-E <robot@dev.null>"
)


function Convert-AuthorToCommitEmail($author)
{
    # workaround for commit authors
    # converts login IDs to email accounts

    $email = ""
    if ($author.IndexOf("@") -eq -1) {

        # remove domain name
        if ($DomainName[0] -ne "") {
            $author = $author.replace($DomainName[0],"")
        }
        # add email, fqdn
        if ($DomainName[1] -ne "") {
            $author = $author + " <" + $author + "@" + $DomainName[1] + ">"
        }
        
    } else {
        $author = $author + " <" + $author + ">"
    }

    $author # return value
}



function Read-PSConfigLastID { # EventTracking
    param(
        $ConfigFile
    )

 if ([string]::IsNullOrEmpty($ConfigFile)) {
        Write-Output "config file name parameter is missing"
        Exit(1)
    }

    $LastID = ""
    # Read config
    #Write-Output "Reading configuration from $ConfigFile"
    Get-Content $ConfigFile | Foreach-Object{
        $var = $_.Split('=')
        if ($var[0] -eq "LastID")
        {
            $LastID = $var[1]
        }
    }
  
    return $LastID
}

function Write-PSConfigLastID { # EventTracking

    param(
        $ConfigFile
        , $LastID        
    )

 if ([string]::IsNullOrEmpty($ConfigFile)) {
        Write-Output "config file name parameter is missing"
        Exit(1)
    }

    Write-Output "Saving configuration (last ID: $LastID) to $ConfigFile"    

    Remove-Item $($ConfigFile+".bak") -ErrorAction SilentlyContinue
    Rename-Item $ConfigFile $($ConfigFile+".bak")
    try {        
        echo LastID=$LastID | Out-File -Append $ConfigFile
    }
    catch {
    Rename-Item $($ConfigFile+".bak") $ConfigFile
    }
}



function Expand-ZipFile($file, $destination)
{
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())   {
        $shell.Namespace($destination).copyhere($item)
        if ($item.Name -like "*%20*") { #rename   
            $OldName = Split-Path -Leaf -Path $item.Path
            $NewName = $OldName -Replace "%20"," "
            $OldName = Join-Path $destination -ChildPath $OldName 
            #Write-Host Renamed $OldName to $NewName
            Rename-Item -path $OldName -newName $NewName -Force
            }
    }
}

function Sync-GIT {
    param(
        $BasePath = $(throw 'base path is required')
        , $GitHome=$HomePath  # PATH TO .ssh FOLDER
    )

    Set-Location $BasePath

    try {
     	$env:HOME=$GitHome
        $output = & git push -u origin master 2>&1
        Write-Output $output
    }
    catch {
        Write-Output "GIT exception: " $_.Exception.Message
    }

    # a workaround for bash: /dev/tty: No such a device or address error - include password
    # git remote -v
    # git remote set-url origin http://username:password@servername:7990/scm/project/repo.git
}

function Publish-GIT {
    param(
        $BasePath
    )
    Write-Output Committing $BasePath now...
    set-location $BasePath
    $author = $GitAuthor
    try {
        git add . -v
    }
    catch {
        Write-Output $_.Exception.Message
    }

    try {
        $msg = "Live Snapshot - "
        $msg += Get-Date -UFormat "%A %m/%d/%Y %R %Z"        
        git commit --author=$author -m "$msg"
    }
    catch {
        Write-Output $_.Exception.Message
    }
}


function Read-PSConfig {

    param(
        $ConfigFile
    )

    if ([string]::IsNullOrEmpty($ConfigFile)) {
        echo "config file name parameter is missing"
        Exit(1)
    }

    $r=0
    # Read config
    Write-Output Reading configuration from $ConfigFile
    Get-Content $ConfigFile | Foreach-Object{
        $var = $_.Split('=')
        New-Variable -Scope Global -Force -Name $var[0] -Value $var[1]
        $r=$r+1
    }

    #if ([string]::IsNullOrEmpty($BasePath) -Or $BasePath.Length -le 5) {
    #    echo "BasePath parameter is wrong or missing in config"
    #    Exit(2)
    #}

    if ($r -eq 0 ) {
        echo "Config file empty - nothing read"
        Exit(2)
    }


}

function Write-PSConfig($ConfigFile,$ConfigKey,$ConfigValue) {

Write-Host "Saving last configuration (Last $ConfigKey -" $ConfigValue ")"
Remove-Item $($ConfigFile+".old") -ErrorAction SilentlyContinue
Rename-Item $ConfigFile $($ConfigFile+".old")
try {
Get-Content $($ConfigFile+".old") | Where-Object {$_ -notlike "$ConfigKey=*"} | Out-File $ConfigFile
echo $ConfigKey=$ConfigValue | Out-File -Append $ConfigFile
}
catch {
Rename-Item $($ConfigFile+".old") $ConfigFile
}
}


function Convert-EventPathFromDDLEvent {

    param(
        $EventType
    )

    $EventPath =  "\StoredProcedure\" ## default
    if ($EventType -like "*_PROCEDURE") { $EventPath = "\StoredProcedure\" }
    else {
        if ($EventType -like "*_PARTITION_FUNCTION") { $EventPath = "\PartitionFunction\" }
        else {
            if ($EventType -like "*_PARTITION_SCHEME") { $EventPath = "\PartitionScheme\" }
            else {
                if ($EventType -like "*_FUNCTION") { $EventPath = "\UserDefinedFunction\" }
                else { 
                    if ($EventType -like "*_TABLE") { $EventPath = "\Table\" }                
                    else { 
                        if ($EventType -like "*_VIEW") { $EventPath = "\View\" }                    
                        else { 
                            $EventPath = ""
                             
                        }
                    }
                }
            }
        }
    }

    return $EventPath

}

function Export-CubeToXML {

    param(
        $cube
        ,$SavePath
        ,$fileExt
    )          
            
            $stringbuilder = new-Object System.Text.StringBuilder
            #$stringwriter = new-Object System.IO.StringWriter($stringbuilder)
            #$xmlOut = New-Object System.Xml.XmlTextWriter($stringwriter)
            $xmlWriter = New-Object System.Xml.XmlTextWriter(new-Object System.IO.StringWriter($stringbuilder))
            $xmlWriter.Formatting = [System.Xml.Formatting]::Indented

            $MSASObject=[Microsoft.AnalysisServices.MajorObject[]] @($cube)
            $scriptObject = New-Object Microsoft.AnalysisServices.Scripter            
            $ScriptObject.ScriptCreate($MSASObject,$xmlWriter,$false)
            
            # Remove CREATE node
            $xmlDoc = New-Object System.Xml.XmlDocument
            $xmlDoc.LoadXml($stringbuilder)            
            $NodeList = $xmlDoc.DocumentElement.Item("ObjectDefinition") #.Item("Cube")

            if ($fileExt -eq ".database") {
                $child = $NodeList.Item("Database")
                Remove-ChildNodeIfExists $child "Cubes"
                Remove-ChildNodeIfExists $child "Dimensions"
                Remove-ChildNodeIfExists $child "DataSources"
                Remove-ChildNodeIfExists $child "DataSourceViews"
                Remove-ChildNodeIfExists $child "Roles"
                Remove-ChildNodeIfExists $child "MiningStructures"

            }

            # finalize
            $stringbuilder = new-Object System.Text.StringBuilder
            $xmlOut = New-Object System.Xml.XmlTextWriter(new-Object System.IO.StringWriter($stringbuilder))
            $xmlOut.Formatting = [System.Xml.Formatting]::Indented
            $NodeList.WriteContentTo($xmlOut)
            
            $filename = $SavePath + "\" + $cube.Name + $fileExt
            #Write-Output 'Scripting' $cube.Name 'to' $filename
            $stringbuilder.ToString() |out-file -Encoding utf8 -filepath $filename


}

function Remove-ChildNodeIfExists {
param (
    $node
    ,$removeNodeName
    )
    
    $child = $node.Item($removeNodeName)
    if ($child) {
        $node.RemoveChild($child)
       }
}


function Remove-FolderPathItems {

    param(
        $FolderPath
    )


    if ([string]::IsNullOrEmpty($FolderPath) -Or $FolderPath.Length -le 5) {
        echo "BasePath parameter is wrong or missing in config"
        Exit(2)
    }

    Write-Output Cleaning up $FolderPath folder...
    get-childitem -Path $FolderPath -recurse -include *.sql | Remove-Item 

}


function Script-SQLDB {
 param(
        $ServerName
        ,$Database
        ,$BasePath
        ,$ConnectionString = $null
    )

    # set-psdebug -strict # catch a few extra bugs
    $ErrorActionPreference = "stop"

    $v = [System.Reflection.Assembly]::LoadWithPartialName( 'Microsoft.SqlServer.SMO')
    if ($v.Location -eq $null) {
        Write-Output "SMO is not installed. See https://learn.microsoft.com/en-us/sql/relational-databases/server-management-objects-smo/installing-smo"
        Write-Output "or run 'Install-Module SqlServer' as an administrator"
        throw "SMO not installed"
        return
    }

    if ((($v.FullName.Split(','))[1].Split('='))[1].Split('.')[0] -ne '9') {
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended') | out-null
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SmoEnum') | out-null
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo') | out-null

    if ($ConnectionString -eq $null) {
        $srv = new-object ("Microsoft.SqlServer.Management.Smo.Server") $ServerName # attach to the server
    } else {
        $conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection
        $conn.ConnectionString = $ConnectionString
        $srv = $conn.Connect()
    }
    if ($srv.ServerType -eq $null) # if it managed to find a server
       {
       Write-Output "Sorry, but I couldn't find Server '$ServerName' "
       return
    }

    Remove-FolderPathItems $BasePath 
    Write-Output Scripting objects...
    $scripter = new-object ("Microsoft.SqlServer.Management.Smo.Scripter") $srv # create the scripter
    $scripter.Options.ToFileOnly = $true
    $scripter.Options.Indexes = $TRUE # add indexes
    # we now get all the object types except extended stored procedures
    # first we get the bitmap of all the object types we want
    $all =[long] [Microsoft.SqlServer.Management.Smo.DatabaseObjectTypes]::all `
        -bxor [Microsoft.SqlServer.Management.Smo.DatabaseObjectTypes]::ExtendedStoredProcedure
    # and we store them in a datatable
    $d = new-object System.Data.Datatable
    # get everything except the servicebroker object, the information schema and system views
    $d=$srv.databases[$Database].EnumObjects([long]0x1FFFFFFF -band $all) | `
        Where-Object {$_.Schema -ne 'sys'-and $_.Schema -ne "information_schema" -and $_.DatabaseObjectTypes -ne 'ServiceBroker' -and $_.DatabaseObjectTypes -ne 'AsymmetricKey' -and $_.DatabaseObjectTypes -ne 'SymmetricKey' -and $_.DatabaseObjectTypes -ne 'SqlAssembly' -and $_.DatabaseObjectTypes -ne 'Certificate'
    }
    if ($($d | measure).Count -le 1 ) {
        Write-Output  "Cannot enumerate objects in the database - permissions problem?";
        return;
        }

    $d | ForEach-Object {  
      $s = $($_.schema -replace '[\\\/\:\.]','-');
      $o = $($_.name -replace '[\\\/\:\.]','-');
      $SavePath=$BasePath + "\" + $_.DatabaseObjectTypes;
      #Write-Output $SavePath


      if (!( Test-Path -path $SavePath )) # create it if not existing
            {Try { New-Item $SavePath -type directory | out-null }
            Catch [system.exception]{
                Write-Error "error while creating '$SavePath' $_"
                return
             }
        }

        if ($s.Length -eq 0) 
	    { 
		    $n = $o,"sql" -join "."	
	    }
	    else
	    {
		    $n = $s,$o,"sql" -join "."
	    }
	    #Write-Output "$SavePath\$n"
        $scripter.Options.Filename = "$SavePath\$n";
    
        # fix encoding    
        $sEnc = [System.Text.Encoding]::UTF8 # System.Text.Encoding.UTF8
        $scripter.Options.Encoding = $sEnc
    
        # Create a single element URN array
        $UrnCollection = new-object ('Microsoft.SqlServer.Management.Smo.urnCollection')
        $URNCollection.add($_.urn)
        # and write out the object to the specified file
        $scripter.script($URNCollection)



     }

    Write-Output All objects scripted.


}


function Export-SQLDBSnapshot {

    param(
        $ServerName
        ,$Database
        ,$BasePath
        ,$ConnectionString = $null
    )

    $BasePath=$BasePath+$Database
    Script-SQLDB -ServerName $ServerName -Database $Database -BasePath $BasePath -ConnectionString $ConnectionString

    Publish-GIT $BasePath
    Sync-GIT $BasePath # git push -u origin master

}


function Export-SQLDBUpdates { # EventTracking

    param(
        $ServerName
        ,$Databases = "" # TBD
        ,$BasePath
        ,$ConnectionString = $null
        ,$ConfigFileName = "LastID.txt"
        ,$LogDbName = "admin"
        ,$LogSqlView = "logs.vw_event_tracking"
    )

  $ConfigFile = Join-Path $BasePath $ConfigFileName
  $LastID = Read-PSConfigLastID $ConfigFile
  if ([string]::IsNullOrEmpty($LastID)) {
        echo "last ID not found in the config file"
        Exit(1)
    }

Write-Output Database list: $Databases   
Write-Output Last processed event ID: $LastID

Write-Output Connecting to $ServerName
$connection = New-Object System.Data.SqlClient.SqlConnection
if ($ConnectionString -eq $null) {
    $ConnectionString = "Server=$ServerName;Database=$LogDbName;Integrated Security=True;Application Name=GIT.PS"
}
$connection.ConnectionString = $ConnectionString
$connection.Open()

$command = $connection.CreateCommand()
$command.CommandText = "SELECT [id]
	  ,event_data.value('(/EVENT_INSTANCE/EventType)[1]', 'varchar(max)') as EventType
	  ,event_data.value('(/EVENT_INSTANCE/PostTime)[1]', 'datetime') as PostTime
	  ,event_data.value('(/EVENT_INSTANCE/LoginName)[1]', 'varchar(max)') as LoginName
	  ,event_data.value('(/EVENT_INSTANCE/DatabaseName)[1]', 'varchar(max)') as DatabaseName
	  ,event_data.value('(/EVENT_INSTANCE/SchemaName)[1]', 'varchar(max)') as SchemaName
	  ,event_data.value('(/EVENT_INSTANCE/ObjectName)[1]', 'varchar(max)') as ObjectName
	  ,event_data.value('(/EVENT_INSTANCE/TSQLCommand/CommandText)[1]', 'varchar(max)') as CommandText
  FROM " + $LogSqlView + " (nolock) where [id]>'" + $LastID + "' ORDER BY ID ASC;"

$result = $command.ExecuteReader()

Write-Output Iterating commits
foreach ($row in $result)
{
    $LastID = $row["ID"]
    echo $LastID 

    $DatabaseName = $($row["DatabaseName"] -replace '[\\\/\:\.]','-')
    # if ($DatabaseName in $Databases) ... skip the databases not on the list

    $EventType = $row["EventType"]
    $EventPath = Convert-EventPathFromDDLEvent($EventType)

    if (! [string]::IsNullOrEmpty($EventPath)) { 
        $pathname = Join-Path (Join-Path $BasePath $DatabaseName) $EventPath
    
        $SchemaName = $($row["SchemaName"] -replace '[\\\/\:\.]','-')
        $ObjectName = $($row["ObjectName"] -replace '[\[\]\\\/\:\.]','-')

        if ($SchemaName.Length -gt 0) {
	     $ObjectName  = $SchemaName + "." + $ObjectName }

        $filename = $ObjectName + ".sql"  
          
        $author = $row["LoginName"]
        $author = Convert-AuthorToCommitEmail($author)
    
        $commitdate = $row["PostTime"]
        $CommandText = $row["CommandText"]
        $message = $EventType + " " + $ObjectName + " @ " + $commitdate + " by " + $author   
        
    
        Write-Output $LastID,$message,$filename
        Write-Output $pathname

        New-Item -ErrorAction SilentlyContinue -type directory -path $pathname
        $filepath = Join-Path $pathname $filename
        $CommandText | Out-File $filepath -Force -Encoding UTF8
        Set-Location $pathname
    
        git add $filename
        git commit --author=$author --date=$commitdate -m $message $filename

    }

    Write-PSConfigLastID $ConfigFile $LastID

}

$connection.Close()
return $LastID

}

function Export-SSISUpdates($ServerName, $BasePath, $TempPath)
{
    $ConfigFile = $BasePath+"ssis.txt"

    if ($ConfigFile) {
        Export-SSIS $ServerName $BasePath $TempPath $ConfigFile
    }

}

function Export-SSISSnapshot
{ param ($ServerName, $BasePath, $TempPath)

    Export-SSIS $ServerName $BasePath $TempPath $null
}

function Export-SSIS
{ param ($ServerName, $BasePath, $TempPath, $ConfigFile = $null, $ConnectionString = $null)
    $BasePath = $BasePath+"SSIS\"
    if ($ConnectionString -eq $null) {
        $ConnectionString = "Data Source=$ServerName;Initial Catalog=SSISDB;Integrated Security=True;Application Name=GIT.PS"; 
    }

    if (!$ConfigFile) { # Snapshot mode - clean the folder
        Write-Output "Cleaning target folder: $BasePath" 
        get-childitem -Path $BasePath -recurse -exclude .git | Remove-Item -recurse
        #Exit(0)
        $ModifiedDate="1/1/1999"

    } else {
        Read-PSConfig $ConfigFile
        if (!$ModifiedDate) {
            Write-Output "Invalid cutoff date"
            Exit(2)
        }
    }

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Last Update: " + $ModifiedDate)

    # get a list of ssis projects
    $sql = "
    SELECT f.[name] as folder_name
          ,pr.[name] as project_name
          ,pr.[object_version_lsn]
	      ,o.created_by
	      ,o.created_time
      FROM [SSISDB].[catalog].[projects] pr
      inner join [SSISDB].[catalog].[folders] f on f.[folder_id] = pr.[folder_id]
      inner join [SSISDB].[internal].[object_versions] o on o.object_version_lsn = pr.object_version_lsn
      WHERE o.created_time > '" + $ModifiedDate + "' order by o.created_time asc;"; 
 
    $con = New-Object Data.SqlClient.SqlConnection; 
    $con.ConnectionString = $ConnectionString
    $con.Open(); 

    $con2 = New-Object Data.SqlClient.SqlConnection; 
    $con2.ConnectionString = $ConnectionString
    $con2.Open(); 

    $cmd = New-Object Data.SqlClient.SqlCommand $sql, $con; 
    $rd = $cmd.ExecuteReader(); 

    While ($rd.Read()) {
    #if ($rd.Read()) {

            $FolderName = $rd.GetString(0) 
            $ProjectName = $rd.GetString(1)
            $author = $rd.GetString(3)
            $Author = Convert-AuthorToCommitEmail($Author)

            $ModifiedDate = $rd[4]
        
            $message = $FolderName + "\" + $ProjectName +" @ " + $ModifiedDate + " by " + $author 
        
            $pathname = Join-Path $BasePath -ChildPath $FolderName
            $pathname = Join-Path $pathname -ChildPath $ProjectName
            $filename = Join-Path $TempPath -ChildPath "\ssis_project.zip"
        
            Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Exporting {0}" -f $pathname); 
            New-Item -Path $pathname -ItemType Directory -ErrorAction SilentlyContinue >$NULL
 
            get-childitem -Path $pathname -recurse -Exclude ".git" | Remove-Item -recurse -Exclude ".git" 


            $sql = "Exec [SSISDB].[catalog].[get_project] @folder_name = '"+$rd.GetString(0)+"' ,@Project_name= '"+$rd.GetString(1)+"';"

            #Write-Host $sql
        
            $cmd2 = New-Object Data.SqlClient.SqlCommand $sql, $con2; 
            $cmd2.CommandTimeout = 300
            $rd2 = $cmd2.ExecuteReader();
            if ($rd2.Read())
            {
                $bt = $rd2.GetSqlBinary(0).Value;
        
                Write-Host $filename

                # New BinaryWriter; existing file will be overwritten. 
                $fs = New-Object System.IO.FileStream ($filename), Create; 
                #$fs = New-Object System.IO.FileStream ($filename), Create, Write; 
                $bw = New-Object System.IO.BinaryWriter($fs); 
 
            
                # Read of complete Blob with GetSqlBinary         
                #$bt = $rd.GetSqlBinary(5).Value; 
                $bw.Write($bt, 0, $bt.Length); 
                $bw.Flush(); 
                $bw.Close(); 
                $fs.Close(); 
            
                Expand-ZipFile $filename $pathname
            
                Set-Location $pathname
                $fn = Get-ChildItem
                if ($fn.Length -gt 0) {

                    if ($ConfigFile) { # commit updates
                        Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Committing updates " + $pathname); 
                        Write-Output $message
                                                     
                        git add .
                        git commit --author="$author" --date=$ModifiedDate -m "$message"
                    }
            
                }


            }    

            $rd2.Close(); 
            $cmd2.Dispose();

    }

    $rd.Close(); 
    $cmd.Dispose(); 
    $con.Close(); 
    $con.Dispose(); 

    if ($ConfigFile) {
        Sync-GIT $BasePath 
        Write-PSConfig $ConfigFile "ModifiedDate" $ModifiedDate
    } else { # snapshot mode
        Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Committing snapshot " + $BasePath); 
        Publish-GIT $BasePath
        Sync-GIT $BasePath
    }

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Finished");

}



function Export-SQLAgent($ServerName, $BasePath)
{

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Started");
    $BasePath=$BasePath+"SQLAgent\"
    if (!(Test-Path $BasePath)) { MKDIR "$BasePath" }

    #Create a new SMO instance 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
     Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Connecting to SMO at " + $ServerName);
    $srv = New-Object "Microsoft.SqlServer.Management.Smo.Server" $ServerName

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + " Cleaning up destination folder: " + $BasePath);
    Get-Childitem -Path $BasePath -recurse -include *.sql | Remove-Item -ErrorAction SilentlyContinue

    #Script out each SQL Server Agent Job for the server
    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Scripting jobs on " + $srv.Name);
    $srv.JobServer.Jobs | foreach-object -process {out-file -Encoding UTF8 -filepath $("$BasePath\" + $($_.Name -replace '\\', '' -replace '\[', '' -replace '\]', '') + ".sql") -inputobject $_.Script() }

    # Test that we have something to commit (any *.sql files)
    $i=0
    Get-ChildItem -Path $BasePath -recurse -include *.sql | foreach-object { $i++}
    if ($i -eq 0) { Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": ERROR - nothing scripted!!!"); }
    else
    {
        Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Committing " + $i + " files...");
        Publish-GIT $BasePath
        Sync-GIT $BasePath # git push -u origin master
    }

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Finished");

}


function Export-SSAS($ServerName,$BasePath,$IncludeDB,$ExcludeDB)
{

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Started");
    $BasePath=$BasePath+"SSAS\"
    if (!(Test-Path $BasePath)) { MKDIR "$BasePath" }

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices") | Out-Null

    Write-Host "Server name: " $ServerName "; excluding databases: " $ExcludeDB
    Write-Host "Including databases: " $IncludeDB

    $server = New-Object Microsoft.AnalysisServices.Server
    $server.connect($ServerName)
    $databases=$server.databases

    Write-Host Cleaning up $BasePath folder...
    get-childitem -Path $BasePath -recurse -include *.xmla,*.cube,*.dim,*.ds,*.dsv,*.role,*.dmm,*.database | Remove-Item -ErrorAction SilentlyContinue >$NULL

    foreach ($database in $databases) {
    
        if ($IncludeDB -like ("*"""+$database.Name+"""*") ) {
    
            Write-Host "Found Database: " $database.Name
            $SavePath = $BasePath + "\" + $database.Name 
            New-Item $SavePath -type directory -ErrorAction SilentlyContinue > $NULL
        
            foreach ($cube in $database.Cubes)  {      # $cube = $database.Cubes[0]
                Export-CubeToXML $cube $SavePath ".cube" }

            foreach ($dim in $database.Dimensions)  {  # $dim = $database.Dimensions[1]
                Export-CubeToXML $dim $SavePath ".dim"  }

            foreach ($ds in $database.DataSources)  {  # $ds = $database.DataSources[0]
                Export-CubeToXML $ds $SavePath ".ds" }

            foreach ($dsv in $database.DataSourceViews)  {  # $dsv = $database.DataSourceViews[0]
                Export-CubeToXML $dsv $SavePath ".dsv" }

            foreach ($role in $database.Roles)  {  # $role = $database.Roles[0]
                Export-CubeToXML $role $SavePath ".role" }

            foreach ($dmm in $database.MiningStructures)  {  # $dmm = $database.MiningStructures[0]
                Export-CubeToXML $dmm $SavePath ".dmm" }
                                
            Export-CubeToXML $database $SavePath ".database" 
        
    #Assemblies                  : {}
    #Accounts                    : {Asset, Balance, Expense, Flow...}
    #DatabasePermissions         : {DatabasePermission}
    #Translations                : {}

     
        
        }
        else { Write-Host "Skipping database: " $database.Name }
    }

    # Test that we have anything to commit
    $i=0
    Get-ChildItem -Path $BasePath -recurse -include *.xmla,*.cube,*.dim,*.ds,*.dsv,*.role,*.dmm,*.database | foreach-object { $i++}

    if ($i -eq 0) { Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": ERROR - nothing scripted!!!"); }
    else
    {
      Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Committing " + $i + " files...");
      Publish-GIT $BasePath
      Sync-GIT $BasePath # git push -u origin master
    }

    Write-Output ((Get-Date -format yyyy-MM-dd-HH:mm:ss) + ": Finished");
}


# Export-Cube-XML
# Export-SSIS-Updates($ServerName, $BasePath, $TempPath, $ModifiedDate)
# Export-SQLDBUpdates
# Export-SQLDBSnapshot
#
#
# git config --global user.name "Wall-E"
# git config --global user.email robot@dev.null

function Export-SQL {
    param(
        $ServerName
        ,$BasePath
        ,$ConnectionString = $null
    )

    Write-Output Connecting to $ServerName
    $connection = New-Object System.Data.SqlClient.SqlConnection
    if ($ConnectionString -eq $null) {    
        $ConnectionString = "Server=$ServerName;Database=master;Integrated Security=True;Application Name=GIT.PS"
    }
    $connection.ConnectionString = $ConnectionString
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = "select name from master.sys.databases where [state] = 0 and name not in ('master','tempdb','model','msdb','DBA','SSISDB') order by name"

    $result = $command.ExecuteReader()

    Write-Output Iterating databases
    foreach ($row in $result)
    {
        $DatabaseName = $row["name"]
        $DatabasePath=$BasePath+$DatabaseName
        New-Item -Path $BasePath -Name $DatabaseName -ItemType "directory"  -ErrorAction SilentlyContinue | Out-Null

        Write-Output Scripting $DatabaseName to $DatabasePath
        Script-SQLDB -ServerName $ServerName -Database $DatabaseName -BasePath $DatabasePath -ConnectionString $ConnectionString

    }


    $connection.Close()

    Publish-GIT $BasePath
    Sync-GIT $BasePath # git push -u origin master

}



function Export-DdlCommits1 { # xe_tsql_ddl
# Prototype (work-in-progress)

    param(
        $ServerName
        ,$RepoBasePath
        ,$SessionName = "xe_tsql_ddl"
        ,[bool]$Commit = $false 
    )

    $Timestamp = Get-Content -Path (Join-Path $RepoBasePath "timestamp.txt")
    $LastTimestamp = [datetime]"2021-01-01"
    if (! [string]::IsNullOrEmpty($Timestamp)) {
        $LastTimestamp = [datetime] $Timestamp
    }
    [datetime]$NewLastTimestamp = $LastTimestamp 
    
    $sql = ";WITH raw_data(t) AS (
    SELECT CONVERT(XML, target_data)
    FROM sys.dm_xe_sessions AS s
    INNER JOIN sys.dm_xe_session_targets AS st
    ON s.[address] = st.event_session_address
    WHERE s.name = '"+$SessionName+"' AND st.target_name = 'ring_buffer'
), xml_data (ed) AS (
    SELECT e.query('.') 
    FROM raw_data 
    CROSS APPLY t.nodes('RingBufferTarget/event') AS x(e)
)

SELECT DatabaseName = DB_NAME(database_id)
, SchemaName = OBJECT_SCHEMA_NAME(object_id, database_id)
, ObjectName = OBJECT_NAME(object_id, database_id)
, EventPath = CASE object_type
	WHEN 'ROLE' THEN 'DatabaseRole'
	WHEN 'PARTITION_FUNCTION' THEN 'PartitionFunction'
	WHEN 'PARTITION_SCHEME' THEN 'PartitionScheme'
	WHEN 'SCHEMA' THEN 'Schema'
	WHEN 'PROC' THEN 'StoredProcedure'
	WHEN 'SYNONYM' THEN 'Synonym'
	WHEN 'TABLE' THEN 'Table'
	WHEN 'USER' THEN 'User'
	WHEN 'FUNCTION' THEN 'UserDefinedFunction'
	WHEN 'VIEW' THEN 'View'
	ELSE '' END
, Author = stuff(login, 1, charindex('\', login), '') + '@' 
	+ iif(charindex('\', login) > 0, substring(login, 1, charindex('\', login) - 1) + '.', '')
	+ 'comcast.com'
, *
FROM (
  SELECT DISTINCT 
    [timestamp]       = ed.value('(event/@timestamp)[1]', 'datetime'),
    [database_id]     = ed.value('(event/data[@name=""database_id""]/value)[1]', 'int'),
    [database_name]   = ed.value('(event/action[@name=""database_name""]/value)[1]', 'nvarchar(128)'),
    [object_type]     = ed.value('(event/data[@name=""object_type""]/text)[1]', 'nvarchar(128)'),
    [object_id]       = ed.value('(event/data[@name=""object_id""]/value)[1]', 'int'),
    [object_name]     = ed.value('(event/data[@name=""object_name""]/value)[1]', 'nvarchar(128)'),
    [session_id]      = ed.value('(event/action[@name=""session_id""]/value)[1]', 'int'),
    [login]           = ed.value('(event/action[@name=""server_principal_name""]/value)[1]', 'nvarchar(128)'),
    [client_hostname] = ed.value('(event/action[@name=""client_hostname""]/value)[1]', 'nvarchar(128)'),
    [client_app_name] = ed.value('(event/action[@name=""client_app_name""]/value)[1]', 'nvarchar(128)'),
    [sql_text]        = ed.value('(event/action[@name=""sql_text""]/value)[1]', 'nvarchar(max)'),
    [phase]           = ed.value('(event/data[@name=""ddl_phase""]/text)[1]',    'nvarchar(128)')
  FROM xml_data
) AS x
WHERE database_id > 4
ORDER BY [timestamp] asc;
"

    $cnnStr = "Database=master;Integrated Security=True;Server=$ServerName;Application Name=GIT.PS"
    Write-Output "Connecting to $ServerName session $SessionName"
    $cnn = New-Object System.Data.SqlClient.SqlConnection($cnnStr)
    $cnn.Open()
    $cmd = $cnn.CreateCommand()
    $cmd.CommandText = $sql
    $cmd.CommandTimeout = 300

    $result = $cmd.ExecuteReader()
    
    Write-Output "Iterating commits since $LastTimestamp"
    foreach ($row in $result)
    {

        $DatabaseName = $($row["DatabaseName"] -replace '[\\\/\:\.]','-')
        $SchemaName = $($row["SchemaName"] -replace '[\\\/\:\.]','-')
        $ObjectName = $($row["ObjectName"] -replace '[\[\]\\\/\:\.]','-')
        $Author = $row["Author"]
        $SqlText = $row["sql_text"]
        $EventPath = $row["EventPath"]

        $CommitDate = [datetime]$row["timestamp"]
        $CommitDate = $CommitDate.AddMilliseconds(-$CommitDate.Millisecond)  ## truncate milliseconds

        if ((! [string]::IsNullOrEmpty($EventPath)) -and (($CommitDate) -gt ($LastTimestamp)) -and ($DatabaseName -ne "DBA") -and (! [string]::IsNullOrEmpty($ObjectName)) )
        { 

            $pathname = Join-Path $RepoBasePath $DatabaseName 
            $pathname = Join-Path $pathname $EventPath 

            if ($SchemaName.Length -gt 0) {
	         $ObjectName  = $SchemaName + "." + $ObjectName }

            $filename = $ObjectName + ".sql"  
          
            # $author = Convert-AuthorToCommitEmail($author)
    
            $message = $ObjectName + " @ " + $CommitDate + " by " + $author   
           
            Write-Output "Committing $message to $pathname\$filename"

            New-Item -ErrorAction SilentlyContinue -type directory -path $pathname | Out-Null
            $SqlText | Out-File $(Join-Path $pathname $filename) -Force -Encoding UTF8
            Set-Location $pathname
    
            if ($Commit) {
                git add $filename
                git commit --author=$author --date=$commitdate -m $message $filename
            }

            if ($CommitDate -gt $NewLastTimestamp) { $NewLastTimestamp = $CommitDate }
        }


    }

    $cnn.Close()
    $NewLastTimestamp | Set-Content -Path (Join-Path $RepoBasePath "timestamp.txt") 
}


function Decode-SQLObjectType {
param ([string]$object_type)

    switch ($object_type) {

        "ROLE" { return "DatabaseRole" }
        "PFUN" { return "PartitionFunction" }
        "PSCHEME" { return "PartitionScheme" }
        "SCHEMA" { return "Schema" }
        "SYNONYM" { return "Synonym" }
        "PROC" { return "StoredProcedure" }
        "USRTAB" { return "Table" }
        "USER" { return "User" }
        "TABFUNC" { return "UserDefinedFunction" }
        "INLFUNC" { return "UserDefinedFunction" }
        "VIEW" { return "View" }
    }

    return ""

}


function Convert-LoginToEmailComcast {
param([string] $login)

    $email = $login 
    $i = 0 + $login.IndexOf("\")
    if ($i -ge 0) { $email = $login.Substring(1+$i) + "@" + $login.Substring(0, $i) + ".comcast.com" }

    return $email 

}


function Export-DdlCommits { # xe_tsql_ddl
param(
    $ServerName
    ,$RepoBasePath
    ,$SessionName = "xe_tsql_ddl"
    ,[bool]$Commit = $false 
)

    Import-Module -Name D:\Tasks\PS\TSQL.psm1 -Force

    ### GET THE LATEST TIMESTAMP
    $Timestamp = Get-Content -Path (Join-Path $RepoBasePath "timestamp.txt")
    $LastTimestamp = [datetime]"2021-01-01"
    if (! [string]::IsNullOrEmpty($Timestamp)) {
        $LastTimestamp = [datetime] $Timestamp
    }
    [datetime]$NewLastTimestamp = $LastTimestamp 


    ### GET LATEST EVENTS
    $sql = "SELECT CONVERT(XML, target_data) xml FROM sys.dm_xe_sessions s INNER JOIN sys.dm_xe_session_targets st ON s.[address] = st.event_session_address
        WHERE s.name = '"+$SessionName+"' AND st.target_name = 'ring_buffer'"
    $cnnStr = "Database=master;Integrated Security=True;Server=$ServerName;Application Name=GIT.PS"
    $cnn = New-Object System.Data.SqlClient.SqlConnection($cnnStr)
    $cnn.Open()
    $cmd = $cnn.CreateCommand()
    $cmd.CommandText = $sql
    $result = $cmd.ExecuteReader()

    if ($result.read()) {

        $count = [int] $xml.RingBufferTarget.eventCount
        Write-Output "Iterating $count commits since $LastTimestamp"

        [xml]$xml = $result["xml"]
        # $content | Set-Content -Path "C:\temp\1d\xml.xml"

        $i=0
        if ($count -gt 0) {
            $xml.RingBufferTarget.event | ForEach-Object {
 
                $i+=1
                $pc = [int](100*$i/($count))
                Write-Progress -Activity "Parsing $count events" -Status "$pc% Complete:" -PercentComplete $pc;
                $timestamp = $_.timestamp
                $action = $_.action
                $data = $_.data

                [int]$database_id = 0
                [int]$object_id = 0
                $object_type = ""
                $object_name = ""
                [string]$sql_text = ""
                [string]$client_login = ""
                [string]$client_hostname = ""

                foreach ($d in $data) {
        
                    if ($d.name -eq "database_id") { $database_id = $d.value }
                    if ($d.name -eq "object_id") { $object_id = $d.value }
                    if ($d.name -eq "object_type") { $object_type = $d.text }
                    if ($d.name -eq "object_name") { $object_name = $d.value }

                }

                foreach ($a in $action) {
        
                    if ($a.name -eq "sql_text") { $sql_text = $a.value }
                    if ($a.name -eq "server_principal_name") { $client_login = $a.value }
                    if ($a.name -eq "client_hostname") { $client_hostname = $a.value }

                }            

                #$DatabaseName = $($row["DatabaseName"] -replace '[\\\/\:\.]','-')
                #$SchemaName = $($row["SchemaName"] -replace '[\\\/\:\.]','-')
                #$ObjectName = $($row["ObjectName"] -replace '[\[\]\\\/\:\.]','-')
                $Author = $client_login + " <"+(Convert-LoginToEmailComcast $client_login)+">"
                $SqlText = $sql_text
                $EventPath = (Decode-SQLObjectType $object_type)

                $CommitDate = [datetime]$timestamp
                $CommitDate = $CommitDate.AddMilliseconds(-$CommitDate.Millisecond)  ## truncate milliseconds

                # exclude tempdb, DBA, and other system dbs
                if (($database_id -ne 2) -and (! [string]::IsNullOrEmpty($EventPath)) `
                    -and ($database_id -gt 5) -and ($object_type -ne 17747)) {
        
                    # obtain actual object name
                    $DatabaseName = [string](Invoke-SQLQueryScalar -CnnStr $cnnStr -SqlStr "SELECT DatabaseName = DB_NAME($database_id)") -replace '[\\\/\:\.]','-'
                    $SchemaName = [string](Invoke-SQLQueryScalar -CnnStr $cnnStr -SqlStr "SELECT SchemaName = OBJECT_SCHEMA_NAME($object_id, $database_id)") -replace '[\\\/\:\.]','-'
                    $ObjectName = [string](Invoke-SQLQueryScalar -CnnStr $cnnStr -SqlStr "SELECT SchemaName = OBJECT_NAME($object_id, $database_id)") -replace '[\[\]\\\/\:\.]','-'

                    if ((! [string]::IsNullOrEmpty($EventPath)) `
                        -and (($CommitDate) -gt ($LastTimestamp)) `
                        -and ($DatabaseName -ne "DBA") `
                        -and (! [string]::IsNullOrEmpty($ObjectName)) )
                    { 

                        $pathname = Join-Path $RepoBasePath $DatabaseName 
                        $pathname = Join-Path $pathname $EventPath 

                        if ($SchemaName.Length -gt 0) { $ObjectName  = $SchemaName + "." + $ObjectName }
                        $filename = $ObjectName + ".sql"  
          
                        $message = $ObjectName + " @ " + $CommitDate + " by " + $author   
           
                        Write-Output "Committing $message to $pathname\$filename"

                        # replace LF with CRLF
                        $SqlText = $SqlText.Replace("`r`n","`n").Replace("`n","`r`n")
                        New-Item -ErrorAction SilentlyContinue -type directory -path $pathname | Out-Null
                        $SqlText | Out-File $(Join-Path $pathname $filename) -Force -Encoding UTF8
                        Set-Location $pathname
    
                        if ($Commit) {
                            git add $filename
                            git commit --author=$author --date=$commitdate -m $message $filename
                        }

                        if ($CommitDate -gt $NewLastTimestamp) { $NewLastTimestamp = $CommitDate }
                    }

                
                }

            } # foreach
        }

    } # read

    $cnn.Close()
    $NewLastTimestamp | Set-Content -Path (Join-Path $RepoBasePath "timestamp.txt") 

}


# Uses DBATOOLS
# Exports multiple DBA Scripts from $items collection into $basepath/$itemtype location
function Export-DBAScripts {
param ($items, $basepath, $itemtype, $CleanFolder = $true)    

    foreach ($item in $items) {

        $filepath = Join-Path (Join-Path $basepath $item.Database) $itemtype
        Mkdir $filepath  -ErrorAction SilentlyContinue | Out-Null 

        if ($CleanFolder -eq $true) {
            Write-Host "Cleaning contents of" $filepath
            Get-ChildItem -Path $filepath -recurse -include *.sql | Remove-Item
            $CleanFolder = $false
        }

        $filename = $($item.Schema -replace '[\\\/\:\.]','-') + "." + $($item.Name -replace '[\\\/\:\.]','-') + ".sql"
        $filename = Join-Path $filepath $filename

        Write-Host $filename 
        $item| Export-DbaScript -Passthru -NoPrefix | Out-File -FilePath $filename -Encoding utf8 -Force
    }
}


# Uses DBATOOLS
# Exports tables, views, stored procedures, udf of all user databases
function Export-SQLDBScripts {
param ($SqlInstance, $SqlCredential, $BasePath, $DatabaseName = $null)

    Import-Module dbatools

    if ($DatabaseName -ne $null) {
        $databases = Get-DbaDatabase -SqlInstance $SqlInstance -SqlCredential $SqlCredential -Database $DatabaseName Normal -Encrypted
    }
    else {
        Write-Host "Retrieving user databases from $SqlInstance"
        $databases = Get-DbaDatabase -SqlInstance $SqlInstance -SqlCredential $SqlCredential -ExcludeSystem -Status Normal -Encrypted
    }

    foreach ($database in $databases) {

        if ($database.IsAccessible) {
            Write-Host "Processing database" $database.Name

            #Table
            $items = Get-DbaDbTable -SqlInstance $SqlInstance -SqlCredential $SqlCredential -Database $database.Name
            Export-DBAScripts $items $BasePath "Table"

            # View
            $items = Get-DbaDbView -SqlInstance $SqlInstance -SqlCredential $SqlCredential -Database $database.Name -ExcludeSystemView
            Export-DBAScripts $items $BasePath "View"

            # StoredProcedure
            $items = Get-DbaDbStoredProcedure -SqlInstance $SqlInstance -SqlCredential $SqlCredential -Database $database.Name -ExcludeSystemSp
            Export-DBAScripts $items $BasePath "StoredProcedure"

            # UserDefinedFunction
            $items = Get-DbaDbUdf -SqlInstance $SqlInstance -SqlCredential $SqlCredential -Database $database.Name -ExcludeSystemUdf
            Export-DBAScripts $items $BasePath "UserDefinedFunction"

        }

    }
}


