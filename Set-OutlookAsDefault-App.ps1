# PowerShell Script to set Outlook as the default app for .ics files using DISM

# Step 1: Export current app associations to an XML file
$exportPath = "C:\Temp\DefaultAppAssociations.xml"
dism /online /Export-DefaultAppAssociations:$exportPath

# Step 2: Load the XML file into PowerShell
[xml]$xmlContent = Get-Content -Path $exportPath

# Step 3: Update the XML content to set Outlook as the default app for .ics files
$icsNode = $xmlContent.DefaultAssociations.Association | Where-Object {$_.ProgId -eq "Outlook.File.ics.15"}
if ($icsNode -eq $null) {
    $newNode = $xmlContent.CreateElement("Association")
    $newNode.SetAttribute("Identifier", ".ics")
    $newNode.SetAttribute("ProgId", "Outlook.File.ics.15")
    $newNode.SetAttribute("ApplicationName", "Microsoft Outlook")
    $xmlContent.DefaultAssociations.AppendChild($newNode)
} else {
    $icsNode.ProgId = "Outlook.File.ics.15"
    $icsNode.ApplicationName = "Microsoft Outlook"
}

# Step 4: Save the modified XML back to disk
$xmlContent.Save($exportPath)

# Step 5: Import the modified XML file to update app associations
dism /online /Import-DefaultAppAssociations:$exportPath

Write-Host "Default app for .ics files set to Microsoft Outlook."
