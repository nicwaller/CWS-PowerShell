function Open-CherwellSession {
<#
.SYNOPSIS
Execute the login action to open a session.
#>
[CmdletBinding()]
param
(
  [Parameter(Mandatory=$True)]
  [String] $Username,
  [Parameter(Mandatory=$True)]
  [String] $Password,
  [Parameter(Mandatory=$True)]
  [String] $Server
)
begin {
  $uri = "https://${server}/cherwellservice/api.asmx"

  $headers = @{SOAPAction="http://cherwellsoftware.com/Login"}
  Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -SessionVariable session -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
<soap:Body>
<Login xmlns="http://cherwellsoftware.com">
<userId>${Username}</userId>
<password>${Password}</password>
</Login>
</soap:Body>
</soap:Envelope>
"@ | Out-Null

# TODO actually check login result instead of assuming a successful authentication

  $sess = New-Object -TypeName PSObject -Prop @{Session=$session;Uri=$uri}
  return $sess
}
} # the end


function Find-CherwellObject {
<#
.SYNOPSIS
Find a business object using an exact-match query. This is NOT a substring match.
There are no wildcards. You need to know exactly what you're looking for, here.
#>
[CmdletBinding()]
param
(
  [Alias('BusOb')]
  [ValidateSet("Problem","Incident","Customer","Task","Service")]
  [String] $BusinessObject = "Incident",

  [Parameter(Mandatory=$True)]
  [Alias('Field')]
  [String] $FieldName,

  [Parameter(Mandatory=$True)]
  [String] $Value,

  [Parameter(Mandatory=$True)]
  $Session
)
begin {
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $Session.session
  $uri = $session.uri

$headers = @{SOAPAction="http://cherwellsoftware.com/QueryByFieldValue"}
$response = Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -WebSession $sess -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <QueryByFieldValue xmlns="http://cherwellsoftware.com">
      <busObNameOrId>${BusinessObject}</busObNameOrId>
      <fieldNameOrId>${FieldName}</fieldNameOrId>
      <value>${Value}</value>
    </QueryByFieldValue>
  </soap:Body>
</soap:Envelope>
"@
$xResponse = [xml] $response
$xChildren = $xResponse | Select-Xml("/*/*/*/*")

$xDeepResponse = [xml] $xChildren.node."#text"
if ($xDeepResponse -eq $null) {
 return $null
}
$xDeepChildren = $xDeepResponse | Select-Xml("/*/*")

$xDeepChildren |
  % {
    New-Object -TypeName PSObject -Prop @{
      'ID'=$_.Node."#text";
      'RecordID'=$_.Node.RecId;
      'TypeID'=$_.Node.TypeId;
    }
  }


}
} #the end


function Find-CherwellObjectUsingQuery {
<#
.SYNOPSIS
Find a business object using a query as defined in the Search Manager.
#>
[CmdletBinding()]
param
(
  [Alias('BusOb')]
  [ValidateSet("Problem","Incident","Customer","Task","Service")]
  [String] $BusinessObject = "Incident",

  [Parameter(Mandatory=$True)]
  [Alias('Query')]
  [String] $QueryName,

  [Parameter(Mandatory=$True)]
  $Session
)
begin {
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $Session.session
  $uri = $session.uri

$headers = @{SOAPAction="http://cherwellsoftware.com/QueryByStoredQuery"}
$response = Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -WebSession $sess -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <QueryByStoredQuery xmlns="http://cherwellsoftware.com">
      <busObNameOrId>${BusinessObject}</busObNameOrId>
      <queryNameOrId>${QueryName}</queryNameOrId>
    </QueryByStoredQuery>
  </soap:Body>
</soap:Envelope>
"@
$xResponse = [xml] $response
$xChildren = $xResponse | Select-Xml("/*/*/*/*")

$xDeepResponse = [xml] $xChildren.node."#text"
if ($xDeepResponse -eq $null) {
 return $null
}
$xDeepChildren = $xDeepResponse | Select-Xml("/*/*")

$xDeepChildren |
  % {
    New-Object -TypeName PSObject -Prop @{
      'ID'=$_.Node."#text";
      'RecordID'=$_.Node.RecId;
      'TypeID'=$_.Node.TypeId;
    }
  }


}
} #the end


function Get-CherwellObject {
<#
.SYNOPSIS
Get exactly one Cherwell business object based on its globally unique record ID.
But you still need to specify what type of object you want.
#>
[CmdletBinding()]
param
(
  [Alias('BusOb')]
  [ValidateSet("Problem","Incident","Customer","Task","Service")]
  [String] $BusinessObject = "Incident",

  [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
  [String] $RecordID,

  [Parameter(Mandatory=$True)]
  $Session
)
begin {
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $Session.session
  $uri = $session.uri
}
process {

$headers = @{SOAPAction="http://cherwellsoftware.com/GetBusinessObject"}
$response = Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -WebSession $sess -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <GetBusinessObject xmlns="http://cherwellsoftware.com">
      <busObNameOrId>${BusinessObject}</busObNameOrId>
      <busObRecId>${RecordID}</busObRecId>
    </GetBusinessObject>

  </soap:Body>
</soap:Envelope>
"@
$xResponse = [xml] $response
$xChildren = $xResponse | Select-Xml("/*/*/*/*")

$xDeepResponse = [xml] $xChildren.node."#text"
if ($xDeepResponse -eq $null) {
 return $null
}
$xDeepChildren = $xDeepResponse | Select-Xml("/*/*")

$r = New-Object -TypeName PSObject
$xDeepResponse |
  Select-Xml("/BusinessObject/FieldList/Field") |
  ForEach-Object {$r | Add-Member -MemberType NoteProperty -Name $_.Node.Name -Value $_.Node."#text"}
$r

}
} #the end



$s = Open-CherwellSession -Username $username -Password $password -Server support.unbc.ca
# TODO should probably close (logout) all the sessions we open
#Find-CherwellObject -BusinessObject "Incident" -FieldName "Status" -Value "Assigned" -Session $s
#Find-CherwellObject -BusinessObject "Incident" -FieldName "Short Description" -Value "prt302 and paper selection" -Session $s
#Find-CherwellObject -FieldName "Status" -Value "New" -Session $s
#Find-CherwellObject -BusinessObject "Incident" -FieldName "Status" -Value "Assigned" -Session $s |

Find-CherwellObjectUsingQuery -Query "My Open Tickets" -Session $s |
  Get-CherwellObject -BusinessObject "Incident" -Session $s |
# Format-Table IncidentID,Service,Category,ShortDescription -Auto
  Select -Expand CreatedDateTime |
  % { ((get-date) - (get-date $_)).totalHours } |
  Measure -Sum |
  Select -Expand Sum
