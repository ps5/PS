<#
.SYNOPSIS
  Various networking cmdlets

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Paul Shiryaev <ps@paulshiryaev.com>
  github.com/ps5
#>

param(
    [parameter(Position=0,Mandatory=$true)][string]$AdminCnnStr # e.g.: 'Persist Security Info=False;Integrated Security=true;Initial Catalog=ADMIN;server="SQLDB-SERVER-NAME"'
)

Import-Module -Name "$PSScriptRoot\Logging.psm1" -Force -ArgumentList $AdminCnnStr 

function Restart-Server {
param ($ComputerName)


    try {
        # $credential = Import-CliXml -Path 'D:\Tasks\PS\ps.xml'
        Restart-Computer -ComputerName $ComputerName -Force -ErrorAction Stop  # -Credential $credential
        $msg = $(Get-Date).ToString() + " restarted $ComputerName"
        Write-AuditLog "X" $msg
    }
    catch {
        $msg = $(Get-Date).ToString() + " failed to restart $ComputerName (" + $PSItem.Exception.Message + ")"
        Write-AuditLog "E" $msg
    }

}

function Get-ServerMemory {
param([string]$ServerName)

    return (Get-CimInstance -Class CIM_PhysicalMemory -ComputerName $ServerName -ErrorAction Stop | Select-Object FormFactor |  Measure-Object FormFactor -Sum).Sum

}

function Test-Http {
param([string]$url)

    try {
        $wr = [System.Net.WebRequest]::Create($url)
        $wr.Method = "GET"
        $wr.UseDefaultCredentials = $true
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        [System.Net.WebResponse] $response = $wr.GetResponse()
        $status = $response.StatusDescription
        if ($response.StatusCode -ne "OK") {
            Write-AuditLog -LogState 'N' -LogMessage "HTTP TEST ERROR $status : $url"

            }
        # $rs = $response.GetResponseStream()
        # [System.IO.StreamReader] $sr = New-Object System.IO.StreamReader -argumentList $rs
        # [string] $results = $sr.ReadToEnd()
        # $results
        }

   catch
    {
        Write-AuditLog -LogState 'E' -LogMessage $_.Exception.Message
    }

}

function Ping-Host {
    param ( [string]$ServerName=$(throw 'server name is required.') )

    try {

        # Check connectivity
        $result = Test-Connection -ComputerName $ServerName -Quiet

        if ($result -eq $false) {

            Write-AuditLog -LogState 'N' -LogMessage "PING FAILED: Cannot reach $ServerName - offline?"
        }
        else {

            # Write-AuditLog -LogState '*' -LogMessage "$ServerName pings okay"
        }
    }
    catch
    {
        Write-AuditLog -LogState 'E' -LogMessage $_.Exception.Message
    }


}

function Test-PortConnect {
    param ( [string]$ServerName=$(throw 'server name is required.')
        , [string]$PortNo=$(throw 'port number is required.') )

    try {

        # Test port open
        $result = New-Object Net.Sockets.TcpClient $ServerName, $PortNo

        if ($result.Connected) {
            # Write-AuditLog -LogState '*' -LogMessage "$ServerName $PortNo connects okay"
        }

    }
    catch
    {
        [string] $msg = $_.Exception.Message
        if ($msg.ToString().Contains("target machine actively refused")) {
            Write-AuditLog -LogState 'N' -LogMessage "CONNECT FAILED: Cannot reach $ServerName port $PortNo - closed?"
            }
        else {
            if ($msg.ToString().Contains("respond after a period of time")) {
            Write-AuditLog -LogState 'N' -LogMessage "CONNECT FAILED: Cannot reach $ServerName port $PortNo - offline?"
            }
            else {
            $msg=$ServerName + $msg.Remove(0, $msg.IndexOf(': ')).Replace('"','')
            Write-AuditLog -LogState 'E' -LogMessage $msg
            }
        }

     }


}



function Test-SQLConnect {
    param ( [string]$ServerName=$(throw 'server name is required.')
        , [string]$PortNo=1433
        , [string]$Database='master'
	, [boolean]$Quiet=$False
    )

    $LogMessage = "$ServerName $PortNo cannot connect"

    try {

        Test-PortConnect $ServerName $PortNo  # Test port open first

        $cnn = New-Object System.Data.SqlClient.SqlConnection
        $cnn.ConnectionString = "Server=$ServerName,$PortNo;Database=$Database;Integrated Security=True;Application Name=NETWORK.PS"
        $cnn.Open()


        if ($cnn.State -eq 1) { # 1 - Open

            $LogMessage = "$ServerName $PortNo connects okay"

            $cmd = $cnn.CreateCommand()
            $cmd.CommandType = [System.Data.CommandType]::Text
            $cmd.CommandText = 'SELECT @@VERSION'
            $result = $cmd.ExecuteReader();
            if ($result.Read()) {
                $LogMessage = $result[0]
            }

        }

        $cnn.Close()


    }
    catch
    {
        [string] $msg = $_.Exception.Message
        $msg=$ServerName + $msg.Remove(0, $msg.IndexOf(': ')).Replace('"','')
        Write-AuditLog -LogState 'E' -LogMessage $msg
        $LogMessage = $msg

     }

    Write-AuditLog -LogState '+' -ProcName $ServerName -LogMessage $LogMessage -NoOutput $Quiet

}




function Test-ServerShares {
param ($server, $verbose = $false)

    try {
        $cs = New-CimSession -ComputerName $server
        $shares = (Get-SmbShare -CimSession $cs)
    }
    catch {
        $msg = $server+": "+$Error[0].ToString()
        Write-AuditLog -LogState 'E' -LogMessage $msg
        # throw $Error
        return
    }

    $c = $shares.Count
    Write-AuditLog -LogState 'S' -LogMessage "Checking $server shares: total of $c"
    $warning=0

    foreach ($share in $shares) {
        if (-not $share.Special) {

            #Write-Output "Checking" $share.Name

            # Check share permissions
            $perms = (Get-SmbShareAccess -Name $share.name -CimSession $cs)
            $violation=0
            foreach ($perm in $perms) {
                $msg = "\\"+$server+"\"+$share.name+" ("+$share.Path+")"

                if ($perm.AccountName -eq "Everyone") {
		            $access=$perm.AccessRight
		            Write-AuditLog -LogState 'W' -LogMessage "Warning - $access access granted to Everyone on $msg"
                    # if ($verbose -ne $false) { Write-Output $msg "- WARNING:"($perm.AccessRight)"access granted to Everyone" }

                    # check filesystem permissions
                    $path = "\\"+$server+"\"+$share.name
                    $acls = @(Get-Acl -Path $path | Select -ExpandProperty Access)
                    foreach ($acl in $acls) {

                        if ($verbose -ne $false) {
				            $msg="$acl.IdentityReference $acl.AccessControlType $acl.FileSystemRights"
				            Write-AuditLog -LogState '*' -LogMessage $msg
				            #Write-Output "* "$acl.IdentityReference $acl.AccessControlType $acl.FileSystemRights
			            }

                        if (($acl.IdentityReference -in @("Everyone","DomainUsers","CABLE\Users","Users")) -and ($acl.AccessControlType -eq "Allow") -and ($perm.AccessRight -eq "Full") -and ($acl.FileSystemRights -contains "FullControl")) {
	                        $msg="VIOLATION DETECTED on $path!!! Access is open to Everyone on both share and NTFS levels !!!"
				            Write-AuditLog -LogState '!' -LogMessage $msg
				            # Write-Output $msg
				            $violation=1
				            $warning=1
                            # auto-remove
                            #if (($perm.AccessRight -eq "Full") -and ($remove -eq "Yes")) {
                            #    $result = (Revoke-SmbShareAccess -CimSession $cs -Name $share.name -AccountName "Everyone" -Force)
                            #    Write-Output "REMOVED: "$result
                            #}

                        }
                    }



                }

            }

            if ($violation -ne 0) {
		exit $violation
		}
        }
    }
    if ($warning -eq "0") { Write-AuditLog -LogState 'F' -LogMessage "OK" }
}



<#
Batch test connectivity to servers in servers.json

Example servers.json:
{"servers": [{
		"server": "SQLSERVER",
		"fqdn": "SERVERNAME.domain.org",
		"group": "PRODUCTION",
		"ports": ["1433"],
		"roles": ["sql"]
	}]
}

Usage:
    Test-Connectivity -ConfigFile "servers.json" -Quiet

#>
function Test-Connectivity {
param ([string]$ConfigFile=$(throw 'servers.json configuration file is required')
,[switch]$Quiet) # suppress output to console

    try {

        $json = (Get-Content $ConfigFile | ConvertFrom-Json)

        foreach ($server in $json.servers) {

            Ping-Host $server.fqdn

            foreach ($port in $server.ports) {
                Test-PortConnect $server.fqdn $port
                }

		    if ("sql" -in $server.roles) {
			Test-SQLConnect $server.fqdn "1433" "master" $Quiet

            	foreach ($db in $server.databases) {
				    Test-SQLConnect $server.fqdn "1433" $db $Quiet
			    }
		    }

        }


        Write-AuditLog -LogState 'F' -LogMessage "Finished successfully"
    }
    catch
    {
        Write-AuditLog -LogState 'E' -LogMessage $_.Exception.Message
    }
}


# Export-ModuleMember *-*
