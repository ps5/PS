    <#
.SYNOPSIS
  JIRA helper cmdlets

.DESCRIPTION
  PS DBA Toolkit

.NOTES
  Created by Paul Shiryaev <ps@paulshiryaev.com> 
  github.com/ps5
#>

param(
    [parameter(Position=0,Mandatory=$true)][string]$username
    , [parameter(Position=1,Mandatory=$true)][string]$password
)

Import-Module -Name "$PSScriptRoot\Passwd.psm1" -Force

### PUBLIC METHODS

function Push-JiraNewTickets {
param (
[String] $sqlQuery = "EXEC Jira.usp_Ticket_GetNew",
[String] $sqlUpdateQuery = "EXEC Jira.usp_Ticket_UpdateIsSentToJira @TicketID, @IssueKey, @IssueID;",
[String] $connectionString = “Server=local;Database=ADMIN;Integrated Security=True;Application Name=JIRA.PS”,
[Bool] $OverrideIssueType = $false
)  


# Get New Tickets
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString + ";Application Name=JIRA.PS"

$connection2 = New-Object System.Data.SqlClient.SqlConnection
$connection2.ConnectionString  = $connection.ConnectionString

$connection.Open()
$command = $connection.CreateCommand()
$connection2.Open()
$command2 = $connection2.CreateCommand()


$command.CommandText = $sqlQuery
$result = $command.ExecuteReader()
if ($result.HasRows) {
    Write-Output "Have new tickets"
    }
else { Write-Output "No new tickets to post"
    }

foreach ($row in $result)
{
    $RecID = $row["TicketID"]  

    [String] $Project = $row["Project"].ToString().TrimEnd()
    # $Project = $Project.ToString().TrimEnd()

    [String] $IssueType = $row["IssueType"].ToString().TrimEnd()  ## Task
    if (!$IssueType) { $IssueType= "Task" }

    [String] $RequestType = $row["RequestType"].ToString().TrimEnd()
    if (!$RequestType) { 
        $RequestType = "" 
    } else { 
        if ($OverrideIssueType) { $IssueType= $RequestType }
    }

    [String] $Summary = $row["Summary"]
    [String] $Description = $row["description"] 
    [String] $Priority = $row["Priority"].ToString().TrimEnd()

    [String] $Component = $row["Component"].ToString().TrimEnd()
    [String] $Reporter = $row["Reporter"].ToString().TrimEnd()
    [String] $Assignee = $row["Assignee"].ToString().TrimEnd()
    [String] $Assignee = $Assignee.Replace("CABLE\\","").Replace("CABLE\","")
    [String] $Filename = $row["FileName"]  
    [String] $target = $row["target"]  

    Write-Host $RecID $Project $Summary

    

    # Post Ticket
    [string] $url = "$target/rest/api/2/issue"
    [string] $data = "{`"fields`":{`"project`":{`"key`":`"$Project`"},`"issuetype`":{`"name`":`"$IssueType`"},`"summary`":`"$Summary`",`"description`":`"$Description`",`"priority`":{`"name`":`"$Priority`"}}}"
    
    Write-Output "POST: " $url $data
    if (1 -eq 1) {

        $response = Invoke-JiraApiX $url $data "POST"
        # Write-Output $response

        $issueID = $response.id
        $issueKey = $response.key
        [String] $newurl = $response.self

    }
    else { # debugging 
        $issueID = 89090285
        $issueKey = "SQI-2080"
        [String] $newurl = "https://tkts.sys.comcast.net/rest/api/2/issue/89090285"    
    }
    Write-Output $issueKey $newurl $issueID
        

    # mark as resolved    
    if (!([String]::IsNullOrEmpty($newurl)))
    {
        # Update reporter, assignee

        if ($Reporter -ne "") {

            $value = "{""name"": """ + $Reporter + """}"
            Update-JiraField $newurl "reporter" $value 
        }

        if ($Assignee -ne "") {

            $value = "{""name"": """ + $Assignee + """}"
            Update-JiraField $newurl "assignee" $value 
        }

        if ($RequestType -ne "") {

            $value = "{""value"": """ + $RequestType + """}"
            Update-JiraField $newurl "customfield_10253" $value 
        }

        if ($Component -ne "") {

            $value = "[{""id"": """ + $Component + """}]"
            Update-JiraField $newurl "components" $value 
        }



        if (!([String]::IsNullOrEmpty($Filename)))
        {
            Write-Output "Have attachment: $Filename"
            $Attachment=$row["Attachment"]
            $ContentType=$row["ContentType"]
            Update-JiraAddAttachment $newurl $Filename $Attachment $ContentType
        }

        # mark ticket as created
        
        Write-Output "Marking TicketID $RecID as created $issueKey (key: $issueID)"

        $command2 = $connection2.CreateCommand()
        $command2.CommandText = $sqlUpdateQuery
        $command2.Parameters.AddWithValue('@TicketID', $RecID) | Out-Null
        $command2.Parameters.AddWithValue('@IssueKey', $issueKey) | Out-Null
        $command2.Parameters.AddWithValue('@IssueID', $issueID) | Out-Null
        
        $rc = $command2.ExecuteNonQuery() 
        

    }
    
    
}

$connection2.Close()
$connection.Close()


}


function Push-JiraNewTicketComments {
param (
[String] $sqlQuery = "EXEC Jira.usp_TicketComment_GetNew",
[String] $sqlUpdateQuery = "EXEC Jira.usp_TicketComment_UpdateIsSentToJira @TicketCommentID;",
[String] $connectionString = “Server=local;Database=ADMIN;Integrated Security=True;Application Name=JIRA.PS”
)  


# Get New Tickets
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString 

$connection2 = New-Object System.Data.SqlClient.SqlConnection
$connection2.ConnectionString  = $connection.ConnectionString 

$connection.Open()
$command = $connection.CreateCommand()
$connection2.Open()
$command2 = $connection2.CreateCommand()


$command.CommandText = $sqlQuery
$result = $command.ExecuteReader()
if ($result.HasRows) {
    Write-Output "Have new ticket comments"
    }

foreach ($row in $result)
{
    $RecID = $row["TicketCommentID"]  

    [string] $issueUrl = $row["IssueURL"] 
    [string] $comment = $row["Comment"] 

    Write-Host $RecID 
        
    # Post Comment
    [string] $url = $issueURL + "/comment"  
    [string] $data = "{`"body`":`"$comment`"}"
    Write-Output "POST: " $url $data
    if (1 -eq 1) {

        $response = Invoke-JiraApiX $url $data "POST"
        $issueID = $response.id
        $issueKey = $response.key
        [String] $newurl = $response.self

    }
    Write-Output $issueKey $newurl $issueID
        
    # mark as resolved    
    if (!([String]::IsNullOrEmpty($newurl)))
    {

        # mark comment as created
        
        Write-Output "Marking TicketCommentID $RecID as created $issueKey (key: $issueID)"

        $command2 = $connection2.CreateCommand()
        $command2.CommandText = $sqlUpdateQuery
        $command2.Parameters.AddWithValue('@TicketCommentID', $RecID) | Out-Null        
        $rc = $command2.ExecuteNonQuery() 
        
    }

    # Update Description    
    [string] $url = $issueURL
    [string] $data = "{ `"update`": { `"description`": [ {`"set`":`"$comment`"} ] } }"
    Write-Output "PUT: " $url $data
    if (1 -eq 1) {
        $response = Invoke-JiraApiX $url $data "PUT"
        $issueID = $response.id
        $issueKey = $response.key
        [String] $newurl = $response.self
    }
    Write-Output $issueKey $newurl $issueID


    
}

$connection2.Close()
$connection.Close()


}


### PRIVATE METHODS


function Update-JiraField
{
param(
$newurl, $field, $value
)
   # Update reporter, assignee
        if ($value -ne "") {
            [String] $data = "{""fields"" : { " `
                + """$field"" : $value" `
                + "}}"

            Write-Output "PUT:" $newurl $data
            $response = Invoke-JiraApiX $newurl $data "PUT"
            Write-Output $response
        }
}

function Update-JiraAddAttachment($newurl, $Filename, $Attachment, $ContentType)
{

    try {
        $filepath="d:\Tasks\JiraETL\tmp\"+$Filename
        $url=$newurl+"/attachments"
        Set-Content -Path $filepath -Value $Attachment -Encoding Byte

        $basicAuth = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$password"))
        $headers = @{
            "Authorization" = $basicAuth
            "X-Atlassian-Token"="nocheck"
        }

        $wc = new-object System.Net.WebClient
        $wc.Headers.Add("Authorization", $headers.Authorization)
        $wc.Headers.Add("X-Atlassian-Token", "nocheck") 
        $x=$wc.UploadFile($url, $filepath)
        $response=[System.Text.Encoding]::ASCII.GetString($x)

        return $response
    }
    catch {
        Write-Warning "Remote Server Response: $($_.Exception.Message)"
        # Write-Warning "Status Code: $($_.Exception.Response.StatusCode)"
        Write-Warning "URI: $url"
        Write-Warning "File: $filepath"
    }

}

function ConvertTo-Base64($string) {
$bytes = [System.Text.Encoding]::UTF8.GetBytes($string);
$encoded = [System.Convert]::ToBase64String($bytes);
return $encoded;
}


  
function Invoke-JiraApiX($url, $body, $method) {
# Write-Warning "Password: $password"
try {
    $basicAuth = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$password"))
    $headers = @{
        "Authorization" = $basicAuth
        "Content-Type"="application/json"
    }
    # $body = ConvertTo-Json $body
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $response = Invoke-RestMethod -Uri $url -Method $method -Headers $headers -Body $body -UseBasicParsing 
    return $response
}
catch {
    Write-Warning "Remote Server Response: $($_.Exception.Message)"
    # Write-Warning "Status Code: $($_.Exception.Response.StatusCode)"
    Write-Warning "URI: $url"
    Write-Warning "Method: $method"
    Write-Warning "Headers: $headers"
    Write-Warning "Body: $body"
}

}

function Invoke-JiraApi($url, $body, $method) {
try {
  
  #Write-Output "Request url: " $url
  #Write-Output "Request body: " $body
  # Write-Warning $password

  $webRequest = [System.Net.WebRequest]::Create($url)
  $webRequest.ContentType = "application/json"
  $BodyStr = [System.Text.Encoding]::UTF8.GetBytes($body)
  $webrequest.ContentLength = $BodyStr.Length
  $webRequest.ServicePoint.Expect100Continue = $false
  $b64 = ConvertTo-Base64($username + ":" + $password);
  $auth = "Basic " + $b64;
  $webRequest.Headers.Add("Authorization", $auth);
  $webRequest.PreAuthenticate = $true
  $webRequest.Method = $method #"POST"
  [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

  $requestStream = $webRequest.GetRequestStream()
  $requestStream.Write($BodyStr, 0, $BodyStr.length)
  $requestStream.Close()
  [System.Net.WebResponse] $resp = $webRequest.GetResponse()
  
  $rs = $resp.GetResponseStream()
  [System.IO.StreamReader] $sr = New-Object System.IO.StreamReader -argumentList $rs
  [string] $results = $sr.ReadToEnd() 
  $issueData = $results | ConvertFrom-Json
  return $issueData

}
  
catch [System.Net.WebException]{
  if ($_.Exception -ne $null -and $_.Exception.Response -ne $null) {
            $errorResult = $_.Exception.Response.GetResponseStream()
            $errorText = (New-Object System.IO.StreamReader($errorResult)).ReadToEnd()
            Write-Warning "The remote server response: $errorText"
            Write-Output $_.Exception.Response.StatusCode
        } else {
            throw $_
        }
  }
}




### PUBLIC METHODS
Export-ModuleMember -Function Push-JiraNewTickets, Push-JiraNewTicketComments, Update-JiraField -Alias *

# Export-ModuleMember *-*
