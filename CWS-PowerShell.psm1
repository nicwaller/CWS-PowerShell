#Set-StrictMode -version 2

function Open-CWSSession {
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
  if ($CWSSession -ne $null) {
    return $CWSSession
  }

  $uri = "https://${server}/cherwellservice/api.asmx"
  $headers = @{SOAPAction="http://cherwellsoftware.com/Login"}

  Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -SessionVariable SOAPSession -Body @"
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
  Set-Variable CWSSession -Scope Global -Value (New-Object -TypeName PSObject -Property @{
    Uri=$uri;
    SOAPSession=$SOAPSession
  })
  return $CWSSession
}
} # the end


function Close-CWSSession {
<#
.SYNOPSIS
Execute the logout action to close a session.
#>
[CmdletBinding()]
param
(
)
begin {
  if ($CWSSession -eq $null) {
    return
  }

  $uri = $CWSSession.Uri
  $headers = @{SOAPAction="http://cherwellsoftware.com/Logout"}

  Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -SessionVariable SOAPSession -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <Logout xmlns="http://cherwellsoftware.com">
  </Logout>
 </soap:Body>
</soap:Envelope>
"@ | Out-Null

# TODO actually check logout result instead of assuming success

  Set-Variable CWSSession -Scope Global -Value $null

}
} # the end


function Find-CWSBusinessObject {
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
  [String] $Value
)
begin {
  if ($CWSSession -eq $null) {
    Write-Error "login!"
  }
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $CWSSession.SOAPSession
  $uri = $CWSSession.Uri

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
      'Type'=$BusinessObject;
      'TypeID'=$_.Node.TypeId;
      'RecID'=$_.Node.RecId;
      'PublicID'=$_.Node."#text";
    }
  }


}
} #the end


function Invoke-CWSStoredQuery {
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
  [String] $QueryName
)
begin {
  if ($CWSSession -eq $null) {
    Write-Error "login!"
  }
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $CWSSession.SOAPSession
  $uri = $CWSSession.uri

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
      'Type'=$BusinessObject;
      'TypeID'=$_.Node.TypeId;
      'RecID'=$_.Node.RecId;
      'PublicID'=$_.Node.TitleText;
    }
  }


}
} #the end


function Get-CWSBusinessObject {
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
  [String] $RecID,

  [Switch]$RefreshCache,

  $RefreshIntervalHours = 3.0
)
begin {
  if ($CWSSession -eq $null) {
    Write-Error "login!"
  }
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $CWSSession.SOAPSession
  $uri = $CWSSession.uri

  $CacheDir = "${env:Temp}\CWSCache"
  if ((get-item "${env:Temp}\CWSCache") -eq $null) {
    New-Item "${env:Temp}\CWSCache" -Type directory
  }
}
process {
  $CacheFile = "${CacheDir}\${RecID}.xml"
  if (Test-Path $CacheFile) {
    $LastUpdate = (Get-Item $CacheFile).LastWriteTime
    $AgeHours = ((Get-Date) - $LastUpdate).TotalHours
    if ($RefreshCache -or ($AgeHours -gt $RefreshIntervalHours)) {
      Remove-Item $CacheFile
    } else {
      return (Import-Clixml $CacheFile)
    }
  }

$headers = @{SOAPAction="http://cherwellsoftware.com/GetBusinessObject"}
$response = Invoke-WebRequest -uri $uri -Method POST -Headers $headers -ContentType "text/xml" -WebSession $sess -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <GetBusinessObject xmlns="http://cherwellsoftware.com">
      <busObNameOrId>${BusinessObject}</busObNameOrId>
      <busObRecId>${RecID}</busObRecId>
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
$r | Export-Clixml $CacheFile
$r

}
} #the end
