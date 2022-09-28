<#
.SYNOPSIS
  Miscellaneous windows auth helper functions

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Paul Shiryaev <ps@paulshiryaev.com> 
  github.com/ps5
#>

[Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime] | Out-Null



function Add-Cred() {
param($Resource,$UserName,$Password)

    $vault = New-Object Windows.Security.Credentials.PasswordVault

    $cred = New-Object windows.Security.Credentials.PasswordCredential
    $cred.Resource = $Resource
    $cred.UserName = $UserName 
    $cred.Password = $Password
    try { $vault.Remove($cred) | Out-Null }
    catch {} # remove existing credential if exists
    $vault.Add($cred)
    Remove-Variable cred # So that we don't have the password lingering in memory!

}


# example: $cnnStr = Get-CredPasswd -Resource "SERVERNAME" -UserName "USERNAME"
function Get-CredPasswd() {
param ($Resource,$UserName)

    $vault = New-Object Windows.Security.Credentials.PasswordVault
    [string]$userPass = ($vault.Retrieve($Resource,$UserName) | Select-Object -First 1).Password
    return $userPass 
}



# deprecated - use Add-Cred() instead
function Set-Passwd {
param ($PasswdFile, $PlainPassword)

$SecurePassword = ConvertTo-SecureString $PlainPassword -AsPlainText -Force
$Encrypted = ConvertFrom-SecureString -SecureString $SecurePassword -Key (1..16)
$Encrypted | Set-Content $PasswdFile 

}

# deprecated - use Get-CredPasswd() instead
function Get-Passwd {
param ([string]$PasswdFile)

$Encrypted = Get-Content $PasswdFile 
$SecurePassword = ConvertTo-SecureString $Encrypted  -Key (1..16)

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
return $UnsecurePassword 

}

Function Test-ADAuthentication {
    param($username,$password)
    
    # Doesn't work anymore:
    # (new-object directoryservices.directoryentry "",$username,$password).psbase.name -ne $null

     
     # Get current domain using logged-on user's credentials
     $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
     $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)

     ($domain.name -ne $null)

    }



# Check existence of a local windows group
function Find-LocalGroup($groupName)
{ 
 return [ADSI]::Exists("WinNT://$Env:COMPUTERNAME/$groupName,group")
}


# Create a local windows group
function Add-LocalGroup($groupName) {
 $groupExist = Find-LocalGroup($groupName)
 if($groupExist -eq $false) {
  Write-Output "Creating group " $groupName
  $Computer = [ADSI]"WinNT://$Env:COMPUTERNAME,Computer"
  $Group = $Computer.Create("Group", $groupName)
  $Group.SetInfo()
  $Group.Description = $groupName
  $Group.SetInfo()
 } else { "Group : $groupName already exist." }
}


# Check for a local windows group on the local machine...
function Find-LocalGroupMember{
param ($groupName,$memberName)
 $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName" 
 $members = @($group.psbase.Invoke("Members"))
 $memberNames = $members | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
 $memberFound = $memberNames -contains $memberName
 return $memberFound
}

# add a user to a windows local group
function Add-LocalGroupMember {
param($groupName, $userName)
 #$group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"
 #$user = [ADSI]"WinNT://$userName"
 $memberExist = Find-LocalGroupMember -groupName $groupName -memberName $userName
 if($memberExist -eq $false) {
  Write-Output "Adding member $userName to group $groupName"
  $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"  
  $user = [ADSI]"WinNT://$userName" 
  $group.Add($user.Path)
 } else { "Member : $userName already exist." }
}


# remove a user from a local windows group
function Remove-LocalGroupMember {
param($groupName, $userName)
 $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"
 $user = [ADSI]"WinNT://$userName"
 $memberExist = Find-LocalGroupMember $groupName $userName
 if($memberExist -eq $true)  {
  Write-Output "Removing member $userName from group $groupName"
  $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"  
  $user = [ADSI]"WinNT://$userName" 
  $group.Remove($user.Path)
 }  else { "Member : $userName not exists." }
}


function Get-LocalGroups {
param(
[string] $MasterServerConnString
, [string] $MasterServerQueryString
)

    [System.Collections.ArrayList]$groupList = @()

    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = $MasterServerConnString;
    $cnn.Open()    

    $cmd = $cnn.CreateCommand()
    $cmd.CommandTimeout = 360
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $MasterServerQueryString
    $result = $cmd.ExecuteReader()

    foreach ($row in $result) {
        $NTGroup = $row[0]
        # $NTLogin = $row[1]
        if (!$groupList.Contains($NTGroup)) { 
            $groupList.Add($NTGroup) > $null
            }
    }

    $cnn.Close()
    return $groupList
}

function Get-LocalGroupMembers {
param($NTGroup
, $ServerName = $Env:COMPUTERNAME)

    # [System.Collections.ArrayList]$membersList = @()

    $membersList = @()

    $group = [ADSI]"WinNT://$ServerName/$NTGroup" 
    $members = @($group.psbase.Invoke("Members"))

    foreach ($member in $members) {

        $memberName = $member | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        # $membersList.Add($memberName) # > $null

        $membersList += $memberName
    }

    return $membersList
}



function Get-MasterMembersList {
param([string]$paramGroupName)

    [System.Collections.ArrayList]$userList = @()

    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = $MasterServerConnString;
    $cnn.Open()    

    $cmd = $cnn.CreateCommand()
    $cmd.CommandTimeout = 360
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = $MasterServerQueryString
    $result = $cmd.ExecuteReader()
    # dataset schema:
    # 0. LocalNtGroupName
    # 1. NTLogin

    foreach ($row in $result) {
        $NTGroup = [string] $row[0]
        $NTLogin = [string] $row[1]
        if ($NTGroup -eq $paramGroupName) {
            if (!$userList.Contains($NTLogin)) { 
                $userList.Add($NTLogin) > $null
            }
        }
    }

    $cnn.Close()
    return $userList
}




function Add-SQLLogin {
param(
[string]$TargetServerConnString
, [string]$groupName
)

    Write-Output "Creating group SQL login $groupName"
    $cnn = New-Object System.Data.SqlClient.SqlConnection
    $cnn.ConnectionString = $TargetServerConnString;
    $cnn.Open()
    $cmd = $cnn.CreateCommand()
    $cmd.CommandType = [System.Data.CommandType]::Text
    $cmd.CommandText = "IF NOT EXISTS (select null from syslogins where name = '$groupName') /* requires securityadmin role membership */
		    CREATE LOGIN [$groupName] FROM WINDOWS WITH DEFAULT_DATABASE=[master];"
    $cmd.ExecuteNonQuery() > $null
    $cnn.Close()
}


function Sync-LocalGroups {
param(
$MasterServerConnString # query source
, $MasterServerQueryString # e.g. 'select LocalNtGroupName, NTLogin from dbo.LoginsToSync' 
, $TargetServerConnString = “Server=$Env:COMPUTERNAME;Database=master;Integrated Security=True;” # local server
)


    $groupList = Get-LocalGroups -MasterServerConnString $MasterServerConnString -MasterServerQueryString $MasterServerQueryString
    foreach ($NTGroup in $groupList) {
        Write-Output "* Processing group $NTGroup"
        Add-LocalGroup $NTGroup
        Add-SQLLogin -TargetServerConnString $TargetServerConnString -groupName "$Env:COMPUTERNAME\$NTGroup"

        [System.Collections.ArrayList] $masterUserList = Get-MasterMembersList $NTGroup

        $membersList = Get-LocalGroupMembers $NTGroup
        [System.Collections.ArrayList] $localUserList = [System.Collections.ArrayList]::new()
        foreach ($NTLogin in $membersList) { [void]$localUserList.Add($NTLogin) }

        foreach($NTLogin in $masterUserList) {
            # make sure to remove domain names from the source query ideally
            # if ($NTLogin -like 'NTDOMAIN*') { $NTLogin = $NTLogin -replace 'NTDOMAIN\\','' }

            try {
                if ((!$localUserList.contains($NTLogin)) -and ($NTLogin)) {
                    Add-LocalGroupMember -groupName $NTGroup -userName $NTLogin 
                    } else {
                    # Write-Output "Already exists"
                    $localUserList.Remove($NTLogin) > $null
                    }
                }
            catch { Write-Output $_.Exception.Message }  
            }

        foreach($userName in $localUserList) {
            try {
                Remove-LocalGroupMember $NTGroup $NTLogin
            }
            catch { Write-Output $_.Exception.Message }  
        }
    } # foreach NTGroup
} 

