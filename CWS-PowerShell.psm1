Set-StrictMode -version 2

function Invoke-CWSRequest {
<#
.SYNOPSIS
Send a SOAP request to execute a function and return the result. Use this AFTER opening a session.
#>
[CmdletBinding()]
param
(
  [Parameter(Mandatory=$True)]
  [String] $Action,

  [System.Collections.Hashtable] $Arguments = @{}
)
begin {
  $SoapTemplate = @'
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <${Action} xmlns="http://cherwellsoftware.com">
   ${ArgStr}
  </${Action}>
 </soap:Body>
</soap:Envelope>
'@

  $uri = $CWSSession.Uri

  #TODO I could even use this for the login part, with a bit of modification.
  if ($CWSSession -eq $null) {
    Write-Error "Must open session before invoking request."
    return $null
  }
  $WebSession = $CWSSession.SOAPSession

  $ArgStr = $Arguments.GetEnumerator() | %{"<{0}>{1}</{0}>`n" -f $_.key,$_.Value}
  $RequestParams = @{
    'WebSession'=$WebSession;
    'Uri'=$Uri;
    'Method'='POST';
    'Headers'=@{SOAPAction="http://cherwellsoftware.com/${Action}"};
    'ContentType'='text/xml';
    'Body'=$ExecutionContext.InvokeCommand.ExpandString($SoapTemplate)
  }

  $Result = Invoke-WebRequest @RequestParams
  Write-Output (([xml]$Result.Content) | Select-Xml("/*/*/*/*")).Node."#text"
}
} # end

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
  $result = Invoke-CWSRequest -Action Logout
  Write-Host -Fore Yellow Logout = $result
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
  $result = Invoke-CWSRequest -Action QueryByFieldValue -Arguments @{'busObNameOrId'=$BusinessObject;'fieldNameOrId'=$FieldName;'value'=$Value}

  $xDeepResponse = [xml] $result
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
  $result = Invoke-CWSRequest -Action QueryByStoredQuery -Arguments @{'busObNameOrId'=$BusinessObject;'queryNameOrId'=$QueryName}

$xDeepResponse = [xml] $result
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

  $result = Invoke-CWSRequest -Action GetBusinessObject -Arguments @{'busObNameOrId'=$BusinessObject;'busObRecId'=$RecID}

$xDeepResponse = [xml] $result
if ($xDeepResponse -eq $null) {
 return $null
}

$r = New-Object -TypeName PSObject
$xDeepResponse |
  Select-Xml("/BusinessObject/FieldList/Field") |
  Where-Object {$_.Node.PSObject.Properties.Match('#text').Count} |
  ForEach-Object {$r | Add-Member -MemberType NoteProperty -Name $_.Node.Name -Value $_.Node."#text" }
$r | Export-Clixml $CacheFile
$r

}
} #the end

function Add-CWSBusinessObjectToMajorIncident {
<#
.SYNOPSIS
Link a business object with a pre-existing Major Incident.
#>
[CmdletBinding()]
param
(
  [Alias('BusOb')]
  [ValidateSet("Problem","Incident","Customer","Task","Service")]
  [String] $BusinessObject = "Incident",

  [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
  [String] $IncidentID,

  [Parameter(Mandatory=$true)]
  [String] $MajorIncidentID
)
begin {
  if ($CWSSession -eq $null) {
    Write-Error "login!"
  }
  [Microsoft.PowerShell.Commands.WebRequestSession] $sess = $CWSSession.SOAPSession
  $uri = $CWSSession.uri

  Add-Type -AssemblyName System.Web
}
process {
  $updateXml = @"
<BusinessObject Name="Incident">
 <FieldList>
  <Field Name="MajorIncidentID">${MajorIncidentID}</Field>
 </FieldList>
</BusinessObject>
"@

  $updateEncoded = [System.Web.HttpUtility]::HtmlEncode($updateXml)

  $result = Invoke-CWSRequest -Action UpdateBusinessObjectByPublicId -Arguments @{'busObNameOrId'=$BusinessObject;'busObPublicId'=$IncidentID;'updateXml'=$updateEncoded}
  $result
}
} #the end
