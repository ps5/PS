<#
.SYNOPSIS
  SFTP helper routines

.DESCRIPTION

.PARAMETER 

.EXAMPLE
  Import-Module -Name "SFTP.psm1" -Force -ArgumentList 'Persist Security Info=False;Integrated Security=true;Initial Catalog=ADMIN;server="$LOGSERVERNAME"'

.NOTES
  Created by Paul Shiryaev, Contractor <ps@paulshiryaev.com>
  github.com/ps5
#>

# set up logging
param([parameter(Position=0,Mandatory=$true)][string]$LogsCnnStr)
Import-Module -Name "$PSScriptRoot\Logging.psm1" -Force -ArgumentList $LogsCnnStr

# set up path to WinSCP .NET DLL assembly here
$PathToWinSCPAssembly = (Join-Path $PSScriptRoot "WinSCPnet.dll")

# internal method
function SftpUpload($localPath,$remotePath,$backupPath, [WinSCP.SessionOptions] $sessionOptions)
{
    try
    {
  
        $session = New-Object WinSCP.Session
        $hostname = $sessionOptions.HostName+","+$sessionOptions.PortNumber
 
        try
        {
            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.FileMask = "|Thumbs.db"
            $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::Off

            $removeFiles = $false

            # Connect
            $dt = (Get-Date).ToString()
            Write-Output "$dt Connecting to $hostname"
            $session.Open($sessionOptions)
            
            if ($session.Opened) {

                if ($remotePath.Substring(-1+$remotePath.Length) -ne "/") { $remotePath += "/" }  ## Add trailing / to the remote path if missing
                Write-Output "Transferring files from $localPath to $remotePath"
                # Upload files, collect results
                $transferResult = $session.PutFiles($localPath+"*", $remotePath, $removeFiles, $transferOptions)
 
                # Iterate over every transfer
                foreach ($transfer in $transferResult.Transfers)
                {
                    # Success or error?
                    if ($transfer.Error -eq $Null)
                    {
                        $subfolder = ($transfer.FileName).Remove(0, $localPath.Length)
                        $filename  = (Join-Path $backupPath $subfolder)
                        $to = Split-Path -Path $filename 
                        $msg="Upload of $($transfer.FileName) succeeded, moving to $filename"
                        # Write-Host $msg
                        Write-AuditLog -LogState 'T' -LogMessage $($transfer.FileName)
                        # Upload succeeded, move source file to backup
                        if(!(Test-Path $to))
                        {
                            New-Item -Path $to -ItemType Directory -Force | Out-Null
                        }
                        if(Test-Path $filename) { Remove-Item $filename -Force | Out-Null }
                        Move-Item $transfer.FileName $to -force #   $backupPath

                    }
                    else
                    {
                        $msg="Upload of $($transfer.FileName) failed: $($transfer.Error.Message)"
                        # Write-Host $msg
                        Write-AuditLog -LogState 'E' -LogMessage $msg
                    }
                }
                Write-AuditLog -LogState 'F' -LogMessage "SFTPUpload ($hostname)"
                } else { Write-AuditLog -LogState 'E' -LogMessage "SFTPUpload error ($hostname) - Cannot connect to the host" }
            }
            finally
            {
                # Disconnect, clean up
                $session.Dispose()
                Write-Output "Done"
                
            }
        
 
        exit 0
    }
    catch [Exception]
    {
        $msg="SFTPUpload exception ($hostname): $($_.Exception.Message)"
        Write-AuditLog -LogState 'E' -LogMessage $msg
        exit 1
    }
}

function Sync-SFTP {
  param ( [string]$UserName=$(throw 'SFTP UserName is required.') 
        , [string]$HostName=$(throw 'SFTP HostName is required.') 
        , [int]$Port=$(throw 'Port number is required.') 
        , [string]$Password='' # use either pasword or key (below) for auth
        , [string]$Key='' # path to ppk private key
        , [string]$RemotePath=$(throw 'Remote SFTP Path is required.') 
        , [string]$Fingerprint=$(throw 'SFTP Fingerprint is required.') 
        , [string]$LocalFilePath = "D:\SFTPSync\"
        , [string]$ArchiveFilePath = "D:\SFTPArchive\"
        )

    # Load WinSCP .NET assembly
    Add-Type -Path $PathToWinSCPAssembly
    #$ScriptPath = $(Split-Path -Parent $MyInvocation.MyCommand.Definition) 
    #[Reflection.Assembly]::LoadFrom( $(Join-Path $ScriptPath "WinSCPnet.dll") ) | Out-Null

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions 
    $sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
    $sessionOptions.HostName = $HostName
    $sessionOptions.PortNumber = [int] $Port
    $sessionOptions.UserName = $UserName
    if ([string]::IsNullOrEmpty($Key)) {
        $sessionOptions.Password = $Password
        } else {
        $sessionOptions.SshPrivateKeyPath = $Key 
        }
    $sessionOptions.SshHostKeyFingerprint = $Fingerprint
    
    $CleansedHostName = (&{If($HostName.Contains(':')) {$HostName.Substring(0,$HostName.IndexOf(':'))} Else {$HostName}})
    $SyncPath =(Join-Path $LocalFilePath ($UserName+"@"+$CleansedHostName))+"\"
    $ArchivePath =(Join-Path $ArchiveFilePath ($UserName+"@"+$CleansedHostName))+"\"

    SftpUpload $SyncPath $RemotePath $ArchivePath $sessionOptions

}



function Get-SFTPFiles {
  param ( [string]$HostName=$(throw 'HostName is required.') 
        , [int]$Port=$(throw 'Port number is required.') 
        , [string]$Fingerprint=$(throw 'Fingerprint is required.') 
        , [string]$UserName=$(throw 'UserName is required.') 
        , [string]$Password=''
        , [string]$Key=''
        , [string]$RemotePath=$(throw 'Remote Path is required.') 
        , [string]$FileMask=$(throw 'Remote Filter is required.') # set to null to download all files and subdirectories of the remote directory.
        , [string]$LocalPath=$(throw 'Local Path is required.') # Full path to the local directory to download the files to.
        , [bool]$NoRemoteFilesFail = $false
        , [bool]$RemoveSourceFileUponTransfer = $false # When set to true, deletes source remote file(s) after a successful transfer. Defaults to false.

        )

    # Load WinSCP .NET assembly
    #Add-Type -Path "WinSCPnet.dll"
    Add-Type -Path (Join-Path $PSScriptRoot "WinSCPnet.dll")
    #$ScriptPath = $(Split-Path -Parent $MyInvocation.MyCommand.Definition) 
    #[Reflection.Assembly]::LoadFrom( $(Join-Path $ScriptPath "WinSCPnet.dll") ) | Out-Null

    $RunTimeError = ""
    $ProcName = "Get-SFTPFiles"
    if ($RemoveSourceFileUponTransfer) { $ProcName += ".Move" } else { $ProcName += "Copy" }

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions 
    $sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
    $sessionOptions.HostName = $HostName
    $sessionOptions.PortNumber = [int] $Port
    $sessionOptions.UserName = $UserName
    if ([string]::IsNullOrEmpty($Key)) {
        $sessionOptions.Password = $Password
        } else {
        $sessionOptions.SshPrivateKeyPath = $Key # "D:\SFTPControl\medallia_comcast_sftp.ppk"
        }
    $sessionOptions.SshHostKeyFingerprint = $Fingerprint
    

    if ($LocalPath.Substring(-1+$LocalPath.Length) -ne "\" `
        -and $LocalPath.Substring(-1+$LocalPath.Length) -ne "*")
            { 
                $LocalPath += "\" }  ## Add trailing \ to the LocalPath if missing

    if ($LocalPath.Substring(-1+$LocalPath.Length) -ne "*") { $LocalPath += "*" }

    # if ($RemotePath.Substring(-1+$RemotePath.Length) -ne "/") { $RemotePath += "/" }  ## Add trailing / to the remote path if missing


    $source = $UserName+"@"+$HostName+$remotePath

    try
    {
  
        $session = New-Object WinSCP.Session
        $hostname = $sessionOptions.HostName+","+$sessionOptions.PortNumber

        try {
            Write-Output $((Get-Date).ToString())" Connecting to $UserName@$HostName"

            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.FileMask = $FileMask # "|Thumbs.db"
            $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::Off

            # Connect
            $session.Open($sessionOptions)
            
            if ($session.Opened) {                

                Write-Output "Transferring $FileMask files from $remotePath to $localPath"
                # Upload files, collect results
                $transferResult = $session.GetFiles($remotePath, $localPath, $RemoveSourceFileUponTransfer, $transferOptions)

                if ($transferResult.IsSuccess) {

                    # Iterate over every transfer
                    foreach ($transfer in $transferResult.Transfers)
                    {

                        # Write-Host ($transfer.FileName)

                        # Success or error?
                        if ($transfer.Error -eq $Null)
                        {
                            $msg=$UserName+"@"+$hostname+$transfer.FileName                        
                            # Write-Host $msg
                            Write-AuditLog -LogState 'T' -LogMessage $msg


                        }
                        else
                        {
                            $msg="Download of $($transfer.FileName) failed: $($transfer.Error.Message)"
                            # Write-Host $msg
                            Write-AuditLog -LogState 'E' -LogMessage $msg
                        }
                    }
                    Write-AuditLog -LogState 'F' -LogMessage "$ProcName ($source)"
                  } else { 
                        if ($NoRemoteFilesFail) {
                           $RunTimeError = "$ProcName ERROR ($source) - Transfer failed" 
                           Write-AuditLog -LogState 'E' -LogMessage $RunTimeError 
                        }
                       }

              } else { 
                    $RunTimeError = "$ProcName ERROR ($source) - Cannot connect to the host" 
                    Write-AuditLog -LogState 'E' -LogMessage $RunTimeError                 
                    }

        }
        finally {
            # Disconnect, clean up
            $session.Dispose()
            Write-Output "Done"
                
        }
        
 
    }
    catch [Exception]
    {
        $RunTimeError="$ProcName EXCEPTION ($source): $($_.Exception.Message)"
        Write-AuditLog -LogState 'E' -LogMessage $RunTimeError
    }


    # throw exception in case of error
    if ([String]::IsNullOrEmpty($RunTimeError)) {
        exit 0
    } else {
        throw $RunTimeError
        exit -1  # stops sqlagent job
    }
}


Export-ModuleMember *-*
