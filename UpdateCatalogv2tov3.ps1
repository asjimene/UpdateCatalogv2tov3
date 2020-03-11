$V3CategoryDefinition = @"
{
    "DisplayName": "",
    "Id": "",
    "Members": [
      ""
    ],
    "ParentId": ""
}
"@


$CatalogType = "Dell"

if ($CatalogType -eq "Dell") {
    $ParentCategories = "Drivers and Applications", "bios", "firmware"
    $ChildCategories = "OptiPlex", "Precision", "Latitude", "XPS", "Alienware", "Inspiron", "Vostro", "PowerEdge"
}


if ($((Get-Item "$PSScriptRoot\*.cab").FullName.Count) -gt 1) {
    $CabFile = (Get-Item "$PSScriptRoot\*.cab" | Out-GridView -Title "Select the CAB to Update" -OutputMode Single).FullName
}
else {
    $CabFile = (Get-Item "$PSScriptRoot\*.cab").FullName
}
$TempCABExtractPath = "$env:TEMP\Catalog-$(New-Guid)"
if (Test-Path $TempCABExtractPath) {
    Remove-Item -Path $TempCABExtractPath -Recurse -Force -Confirm:$false
}
New-Item -ItemType Directory -Path $TempCABExtractPath -Force
Write-Output "Extracting Cab: $Cabfile"
Start-Process expand.exe -ArgumentList "`"$CABFile`" -F:* `"$TempCABExtractPath`"" -Wait
Write-Output "Completed Extracting Cab file"

Write-Output "importing Catalog"
$CatalogXML = [xml](Get-Content "$TempCABExtractPath\*.xml")
Write-Output "Completed Catalog Import"

New-Item -ItemType Directory -Path $TempCABExtractPath -Name "V3" -Force -ErrorAction SilentlyContinue
$AllCategoryIDs = @()

Write-Output "Processing Manufacturer Category: $CatalogType"
$ManufacturerCategoryID = (New-Guid).ToString()
$AllCategoryIDs += $ManufacturerCategoryID
$ManufacturerCategoryMembers = $CatalogXML.SystemsManagementCatalog.SoftwareDistributionPackage

$ManufacturerJson = $V3CategoryDefinition | ConvertFrom-Json
$ManufacturerJson.DisplayName = $CatalogType
$ManufacturerJson.Id = $ManufacturerCategoryID
$ManufacturerJson.Members = $ManufacturerCategoryMembers.Properties.PackageID
$ManufacturerJson | ConvertTo-Json | Out-File -FilePath "$TempCABExtractPath\V3\$ManufacturerCategoryID.json" -Encoding utf8 -Force
foreach ($ParentCategory in $ParentCategories) {
    Write-Output "Processing Parent Category: $ParentCategory"
    $ParentCategoryID = (New-Guid).ToString()
    $AllCategoryIDs += $ParentCategoryID
    $ParentCategoryMembers = $CatalogXML.SystemsManagementCatalog.SoftwareDistributionPackage | Where-Object { $_.Properties.ProductName -eq $ParentCategory }

    $ParentJson = $V3CategoryDefinition | ConvertFrom-Json
    $ParentJson.DisplayName = $ParentCategory
    $ParentJson.Id = $ParentCategoryID
    $ParentJson.Members = $ParentCategoryMembers.Properties.PackageID
    $ParentJson.ParentId = $ManufacturerCategoryID
    $ParentJson | ConvertTo-Json | Out-File -FilePath "$TempCABExtractPath\V3\$ParentCategoryID.json" -Encoding utf8 -Force
    foreach ($ChildCategory in $ChildCategories) {
        Write-Output "Processing Child Category: $ChildCategory"
        $ChildCategoryID = (New-Guid).ToString()
        $AllCategoryIDs += $ChildCategoryID
        $ChildCategoryMembers = $ParentCategoryMembers | Where-Object { $_.LocalizedProperties.Description -match $ChildCategory }
        
        $ChildJson = $V3CategoryDefinition | ConvertFrom-Json
        $ChildJson.DisplayName = "$ParentCategory - $ChildCategory"
        $ChildJson.Id = $ChildCategoryID
        $ChildJson.Members = $ChildCategoryMembers.Properties.PackageID
        $ChildJson.ParentID = $ManufacturerCategoryID
        $ChildJson | ConvertTo-Json | Out-File -FilePath "$TempCABExtractPath\V3\$ChildCategoryID.json" -Encoding utf8 -Force
    }
}

$AllCategoryIds | ConvertTo-Json | Out-File -FilePath "$TempCABExtractPath\V3\update_categories.json" -Force

$MakeCabFileContents = @"
.OPTION EXPLICIT
.Set CabinetNameTemplate=$(($Cabfile.Split("\")[-1]).replace(".cab","-v3.cab"))
.set DiskDirectoryTemplate=$PSScriptRoot
.Set CompressionType=MSZIP
.Set MaxDiskSize=0
.Set RptFileName=$PSScriptRoot\Logs\MakeCABReport$(Get-Date -f FileDateTime).log
.Set Cabinet=on
.Set Compress=on`r`n
"@

Get-ChildItem -Path $TempCABExtractPath -Depth 0 | Sort-Object Attributes -Descending | ForEach-Object {
    if (-not ($_.Attributes -contains "directory")) {
        $MakeCabFileContents += "`"$($_.FullName)`""
    }
    else {
        $MakeCabFileContents += ".Set DestinationDir=$($_.Name)`r`n`""
        $MakeCabFileContents += ((Get-ChildItem $($_.FullName) | Where-Object -Property Attributes -NotContains "directory").Fullname -join "`"`r`n`"")
        $MakeCabFileContents += '"'
    }
    $MakeCabFileContents += "`r`n"
}

$MakeCabFileContents | Out-File $PSScriptRoot\MakecabDirectives.txt -Encoding utf8 -Force
$MakecabProc = Start-Process makecab.exe -ArgumentList "/f", "`"$("$PSScriptRoot\MakecabDirectives.txt")`"" -NoNewWindow -PassThru

## Cleanup
While ($MakecabProc.HasExited -eq $false) {
    Start-Sleep 1
}