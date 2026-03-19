<#
Add-OrUpdate-MDMetadata.ps1

Purpose
-------
• If a Markdown file has NO YAML front matter → insert a metadata block.
• If a Markdown file already has YAML → append missing fields only.
• Existing YAML structure (lists, arrays, comments) is preserved.
• Prevents double YAML blocks.
• Works safely with Obsidian vaults.

PowerShell
----------
Compatible with Windows PowerShell 5.1
#>

[CmdletBinding()]
param(
    [string]$Path = (Get-Location).Path,
    [string]$Author = "Ryan Robichaud",
    [string]$Email = "ruggedroofing777@gmail.com",
    [string]$DOB = "1987-02-12",
    [string]$License = "All Rights Reserved",
    [string]$CopyrightTemplate = "© {0} {1}",
    [switch]$Recurse = $true,
    [switch]$WhatIf = $false
)

Set-StrictMode -Version 2

# Ensure $Path always points to the current working directory if it was empty
if ([string]::IsNullOrWhiteSpace($Path)) {
    $Path = (Get-Location).Path
}

Write-Host "Using path: $Path" -ForegroundColor Green

Set-StrictMode -Version 2

#-----------------------------------
# Remove leading fake YAML delimiter
# This field safely removes "---" if it is the first non-empty line and is 
# followed by a comment line, which is a common pattern when users accidentally 
# paste YAML content without realizing it creates a fake front matter block.
# Takes into account empty rows and only removes the delimiter if it is truly a 
# leading artifact, not part of valid YAML.
#---------------------------------------------------------

function Remove-FakeYamlDelimiter {
    param([string]$Text)

    $lines = $Text -split "`r?`n"

    $firstContentIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -ne "") {
            $firstContentIndex = $i
            break
        }
    }

    if ($firstContentIndex -eq -1) { return $Text }

    $firstLine  = $lines[$firstContentIndex].Trim()
    $secondLine = ($lines[($firstContentIndex + 1)..($lines.Count - 1)] | Where-Object { $_.Trim() -ne "" })[0].Trim()

    if ($firstLine -eq '---' -and $secondLine -match '^#') {
        $newLines = @()
        if ($firstContentIndex -gt 0) {
            $newLines += $lines[0..($firstContentIndex - 1)]
        }
        $newLines += $lines[($firstContentIndex + 1)..($lines.Count - 1)]
        return ($newLines -join "`n")
    }

    return $Text
}

#-----------------------------------
# Main fix pass for leading delimiter

if (-not (Test-Path $Path)) { throw "Path not found: $Path" }

$files = if ($Recurse) { Get-ChildItem $Path -Filter "*.md" -File -Recurse } else { Get-ChildItem $Path -Filter "*.md" -File }

foreach ($file in $files) {
    try {
        $content = Read-TextFile $file.FullName
        $newContent = Remove-FakeYamlDelimiter $content

        if ($newContent -ne $content) {
            if ($WhatIf) {
                Write-Host "[WHATIF] Would remove leading '---' from: $($file.FullName)"
            } else {
                Write-TextFileUtf8NoBom $file.FullName $newContent
                Write-Host "[FIXED] Removed leading '---' from: $($file.FullName)"
            }
        } else {
            Write-Host "[SKIPPED] No change: $($file.FullName)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "[ERROR] $($file.FullName)" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

#-----------------------------------
# Encoding helpers
function Get-Utf8NoBom { New-Object System.Text.UTF8Encoding($false) }

function Read-TextFile {
    param([string]$FilePath)
    $sr = New-Object System.IO.StreamReader($FilePath,$true)
    try { $sr.ReadToEnd() }
    finally {
        $sr.Close()
        $sr.Dispose()
    }
}

function Write-TextFileUtf8NoBom {
    param([string]$FilePath, [string]$Content)
    [System.IO.File]::WriteAllText($FilePath, $Content, (Get-Utf8NoBom))
}

function Get-NewlineStyle {
    param([string]$Text)
    if ($Text -match "`r`n") { return "`r`n" }
    if ($Text -match "`n") { return "`n" }
    return [Environment]::NewLine
}

#-----------------------------------
# Field specification objects
function New-ScalarFieldSpec { param([string]$Name,[string]$Value) [pscustomobject]@{ Name=$Name; Type="Scalar"; Value=$Value } }
function New-BareFieldSpec { param([string]$Name) [pscustomobject]@{ Name=$Name; Type="Bare" } }

function Build-FieldSpecs {
    param($AuthorValue,$EmailValue,$DOBValue,$LicenseValue,$CopyrightValue,$CreatedValue,$EditedValue)

# Edit this section to customize what new fields gets added to YAML Frontmatter
# Also determines order of new and existing fields when inserting new YAML block

    $specs = New-Object System.Collections.Generic.List[object]
    [void]$specs.Add((New-BareFieldSpec "identity"))
    [void]$specs.Add((New-ScalarFieldSpec "category" ""))
    [void]$specs.Add((New-ScalarFieldSpec "language" ""))
    [void]$specs.Add((New-BareFieldSpec "topic"))
    [void]$specs.Add((New-ScalarFieldSpec "status" ""))
    [void]$specs.Add((New-ScalarFieldSpec "created" $CreatedValue))
    [void]$specs.Add((New-ScalarFieldSpec "edited" $EditedValue))
    [void]$specs.Add((New-ScalarFieldSpec "priority" ""))
    [void]$specs.Add((New-ScalarFieldSpec "backedup" "yes"))
    [void]$specs.Add((New-ScalarFieldSpec "hash_sha256" ""))
    [void]$specs.Add((New-ScalarFieldSpec "hash_date" ""))
    [void]$specs.Add((New-ScalarFieldSpec "version" ""))
    [void]$specs.Add((New-ScalarFieldSpec "author" $AuthorValue))
    [void]$specs.Add((New-ScalarFieldSpec "dob" $DOBValue))
    [void]$specs.Add((New-ScalarFieldSpec "email" $EmailValue))
    [void]$specs.Add((New-ScalarFieldSpec "license" $LicenseValue))
    [void]$specs.Add((New-ScalarFieldSpec "copyright" $CopyrightValue))
    return $specs
}

#-----------------------------------
# Frontmatter detection
function Get-FrontMatter {
    param([string]$Text)

    $result = [ordered]@{ HasFrontMatter=$false; FrontMatterRaw=""; Body=$Text; StartDelimiter=""; EndDelimiter="" }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $result }

    $lines = $Text -split "`r?`n"
    $firstContentIndex = -1
    for ($i=0; $i -lt $lines.Count; $i++) { if (-not [string]::IsNullOrWhiteSpace($lines[$i])) { $firstContentIndex=$i; break } }
    if ($firstContentIndex -eq -1) { return $result }
    if (($lines[$firstContentIndex]).Trim() -notmatch '^(---|\.\.\.)\s*$') { return $result }

    $startDelimiter = ($lines[$firstContentIndex]).Trim()
    $endIndex = -1
    for ($i = $firstContentIndex+1; $i -lt $lines.Count; $i++) { if (($lines[$i]).Trim() -eq $startDelimiter) { $endIndex=$i; break } }
    if ($endIndex -eq -1) { return $result }

    $yamlLines = $lines[($firstContentIndex+1)..($endIndex-1)]
    $bodyLines = if ($endIndex+1 -le $lines.Count-1) { $lines[($endIndex+1)..($lines.Count-1)] } else { @() }

    $result.HasFrontMatter = $true
    $result.FrontMatterRaw = ($yamlLines -join "`n")
    $result.Body = ($bodyLines -join "`n")
    $result.StartDelimiter = $startDelimiter
    $result.EndDelimiter = $startDelimiter

    return $result
}

#-----------------------------------
# YAML key detection
function Get-ExistingKeys {
    param([string]$Yaml)
    $keys=@{}
    foreach ($line in ($Yaml -split "`r?`n")) { if ($line -match '^\s*([A-Za-z0-9_.-]+)\s*:') { $keys[$matches[1].ToLower()]=$true } }
    return $keys
}

#-----------------------------------
# Append missing YAML fields
function Add-MissingFields {
    param($Yaml,$CanonicalSpecs,$ExistingKeys,$Newline)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.AddRange($Yaml -split "`r?`n")
    foreach ($spec in $CanonicalSpecs) {
        $key=$spec.Name.ToLower()
        if (-not $ExistingKeys.ContainsKey($key)) {
            if ($spec.Type -eq "Scalar") { $lines.Add("$($spec.Name): `"$($spec.Value)`"") } else { $lines.Add("$($spec.Name):") }
        }
    }
    return ($lines -join $Newline)
}

#-----------------------------------
# Build fresh YAML
function Build-NewYaml {
    param($CanonicalSpecs,$Newline)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($spec in $CanonicalSpecs) {
        if ($spec.Type -eq "Scalar") { $lines.Add("$($spec.Name): `"$($spec.Value)`"") } else { $lines.Add("$($spec.Name):") }
    }
    return ($lines -join $Newline)
}

#-----------------------------------
# Process single markdown file
function Process-MarkdownFile {
    param([string]$FilePath,[object[]]$DesiredFieldSpecs,[switch]$WhatIfMode)
    $content = Read-TextFile $FilePath
    $newline = Get-NewlineStyle $content
    $parsed  = Get-FrontMatter $content

    if ($parsed.HasFrontMatter) {
        $existingKeys = Get-ExistingKeys $parsed.FrontMatterRaw
        $newYaml = Add-MissingFields $parsed.FrontMatterRaw $DesiredFieldSpecs $existingKeys $newline
        $rebuilt = $parsed.StartDelimiter + $newline + $newYaml + $newline + $parsed.EndDelimiter + $newline + $parsed.Body.TrimStart()
        $action = "Updated"
    } else {
        $newYaml = Build-NewYaml $DesiredFieldSpecs $newline
        $rebuilt = "---" + $newline + $newYaml + $newline + "---" + $newline + $content.TrimStart()
        $action = "Inserted"
    }

    if (-not $WhatIfMode) { Write-TextFileUtf8NoBom $FilePath $rebuilt }

    [pscustomobject]@{ File=$FilePath; Action=$action }
}

#-----------------------------------
# Script execution
if (-not (Test-Path $Path)) { throw "Path not found: $Path" }

$today = Get-Date -Format "yyyy-MM-dd"
$year  = Get-Date -Format "yyyy"
$copyrightValue = [string]::Format($CopyrightTemplate,$year,$Author)
$desiredFieldSpecs = Build-FieldSpecs $Author $Email $DOB $License $copyrightValue $today $today

Write-Host ""
Write-Host "Markdown metadata ownership pass" -ForegroundColor Cyan
Write-Host "Path: $Path" -ForegroundColor Green
Write-Host ""

$files = if ($Recurse) { Get-ChildItem $Path -Filter "*.md" -File -Recurse } else { Get-ChildItem $Path -Filter "*.md" -File }

foreach ($file in $files) {
    try {
        $result = Process-MarkdownFile $file.FullName $desiredFieldSpecs -WhatIfMode:$WhatIf
        Write-Host "[$($result.Action)] $($result.File)"
    } catch {
        Write-Host "[ERROR] $($file.FullName)" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

#-----------------------------------
# Footer Display
Write-Host @"
██████╗ ██╗   ██╗ █████╗  ███╗   ██╗
██╔══██╗╚██╗ ██╔╝██╔══██╗ ████╗  ██║
██████╔╝ ╚████╔╝ ███████║ ██╔██╗ ██║
██╔══██╗  ╚██╔╝  ██╔══██║ ██║╚██╗██║
██║  ██║   ██║   ██║  ██║ ██║ ╚████║
╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝ ╚═╝  ╚═══╝
🏖️🌴🌴🌴😎😎😎 All Currencies Accepted‼️😎😎😎🌴🌴🌴🏖️
"@ -ForegroundColor Green

Write-Host "╔═════════════════════════════════════════════════════╗" -ForegroundColor Darkcyan
Write-Host "║  Author: Ryan Robichaud                             ║" -ForegroundColor Darkcyan
Write-Host "║  D.O.B. 1987-02-12                                  ║" -ForegroundColor Darkcyan
Write-Host "║  E-mail: ruggedroofing777@gmail.com                 ║" -ForegroundColor Darkcyan
Write-Host "║  License: All Rights Reserved                       ║" -ForegroundColor Darkcyan
Write-Host "╚═════════════════════════════════════════════════════╝" -ForegroundColor Darkcyan