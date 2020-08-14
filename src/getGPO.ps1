Get-GPOReport -All -ReportType XML -path "a.xml"
$doc = [xml](Get-Content "a.xml") 
$osName=$doc.createNode("element","OSName","")
$osName.InnerText=systeminfo | findstr /B /C:"OS Name"
$osName.InnerText=$osName.InnerText -replace 'OS Name:',''
$osName.InnerText.Trim()
$osVersion=$doc.createNode("element","OSVersion","")
$osVersion.InnerText=systeminfo | findstr /B /C:"OS Version"
$osVersion.InnerText=$osVersion.InnerText -replace 'OS Version:',''
$osVersion.InnerText.Trim()
$hostName=$doc.createNode("element","hostName","")
$hostName.InnerText=systeminfo | findstr /B /C:"Host Name"
$hostName.InnerText=$hostName.InnerText -replace 'Host Name:',''
$hostName.InnerText=$hostName.InnerText -replace '-',''
$hostName.InnerText=$hostName.InnerText.Trim()
#$hostName.InnerText=$hostName.InnerText -replace ' ','_'

$doc.documentElement.appendchild($osName)
$doc.documentElement.appendchild($osVersion)
$doc.documentElement.appendchild($hostName)
$target=$Env:HOMEPATH + "\"+"GPO_"+$hostName.InnerText+".xml"
$doc.save($target)

