$xmlPath = "C:\Users\markstan\Documents\GitHub\IntuneOneDataCollector\Intune.XML"
[xml]$IntuneXML = Get-Content $xmlPath

$regTable =   $IntuneXML.DataPoints.Package.Registries.Registry          | Select-Object @{Name="Output File Path"; Expression="OutputFileName"}, @{Name="Path"; Expression="`#text"} | ConvertTo-Html -Fragment
$fileTable =  $IntuneXML.DataPoints.Package.Files.File | Sort Team       | Select-Object @{Name="Classification"; Expression="Team"}, @{Name="Path"; Expression="`#text"} | ConvertTo-Html -Fragment
$eventTable = $IntuneXML.DataPoints.Package.EventLogs.EventLog           | Select-Object @{Name="Classification"; Expression="Team"}, @{Name="Path"; Expression="`#text"} | ConvertTo-Html -Fragment
$commands =   $IntuneXML.DataPoints.Package.Commands.Command    | Sort Team | Select-Object @{Name="Classification"; Expression="Team"}, @{Name="Output File"; Expression="OutputFileName"}, @{Name="Command"; Expression="`#text"} | ConvertTo-Html -Fragment   | % { $_ -replace "`n", '<br>' }

$header = @"
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Intune One Data Collector</title>

<style type="text/css">


TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;margin-bottom: 50px }
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
tr:nth-child(even) {background-color: #f2f2f2;}
H2 {
color: white;
background-color: #6495ED;
font-family: arial, sans-serif;
font-size: 24px;
font-weight: bold;
margin-top: 0px;
margin-bottom: 20px;
text-align:center;}
</style>
</head><body>

"@

$footer = "</body></html>"

$header     | Out-File dataCollected.html
"<H2>Registry keys</H2>"| Out-File dataCollected.html -Append
$regTable   | Out-File dataCollected.html -Append

"<H2>Files</H2>"| Out-File dataCollected.html -Append
$fileTable  | Out-File dataCollected.html -Append

"<H2>Event Logs</H2>"| Out-File dataCollected.html -Append
$eventTable | Out-File dataCollected.html -Append

"<H2>Commands</H2>"| Out-File dataCollected.html -Append
$commands   | Out-File dataCollected.html -Append

$footer     | Out-File dataCollected.html -Append

start dataCollected.html