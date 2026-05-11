$ErrorActionPreference = "Stop"

# Dynamisk sti til skrivebordet (Tvinges som rå tekststreng)
$desktopPath = [Environment]::GetFolderPath("Desktop").ToString()
$outputPath = "$desktopPath\member-programs.json"


$sources = @(
  @{
    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    Source = "HKLM64"
  },
  @{
    Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    Source = "HKLM32"
  },
  @{
    Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    Source = "HKCU"
  }
)

function Clean-String {
  param([object]$Value)

  if ($null -eq $Value) {
    return $null
  }

  $text = [string]$Value
  $text = $text.Trim()

  if ([string]::IsNullOrWhiteSpace($text)) {
    return $null
  }

  return $text
}

function Limit-String {
  param(
    [object]$Value,
    [int]$MaxLength
  )

  $text = Clean-String $Value

  if ($null -eq $text) {
    return $null
  }

  if ($text.Length -gt $MaxLength) {
    return $text.Substring(0, $MaxLength)
  }

  return $text
}

function Test-ValidInstallDate {
  param([object]$Value)

  $text = Clean-String $Value

  if ($null -eq $text) {
    return $null
  }

  if ($text -notmatch '^\d{8}$') {
    return $text
  }

  try {
    [datetime]::ParseExact($text, "yyyyMMdd", $null) | Out-Null
    return $text
  }
  catch {
    # Bevar rå værdi, men marker den ikke som parsed dato.
    return $text
  }
}

function Test-NoiseProgram {
  param(
    [string]$Name,
    [object]$SystemComponent,
    [object]$ReleaseType,
    [object]$ParentKeyName
  )

  if ($SystemComponent -eq 1) {
    return $true
  }

  if (-not [string]::IsNullOrWhiteSpace($ParentKeyName)) {
    return $true
  }

  if ($ReleaseType -in @("Update", "Hotfix", "Security Update")) {
    return $true
  }

  return $false
}

$programs = foreach ($source in $sources) {
  Get-ItemProperty $source.Path -ErrorAction SilentlyContinue |
    Where-Object {
      $_.DisplayName -and
      -not (Test-NoiseProgram `
        -Name $_.DisplayName `
        -SystemComponent $_.SystemComponent `
        -ReleaseType $_.ReleaseType `
        -ParentKeyName $_.ParentKeyName)
    } |
    ForEach-Object {
      [PSCustomObject]@{
        name         = Limit-String $_.DisplayName 255
        version      = Limit-String $_.DisplayVersion 100
        publisher    = Limit-String $_.Publisher 255
        install_date = Limit-String (Test-ValidInstallDate $_.InstallDate) 50
        registry_key = Limit-String $_.PSChildName 255
        source       = $source.Source
      }
    }
}

# Dedup på samme nøgle som Prisma unique constraint.
$deduped = $programs |
  Where-Object { $_.name } |
  Sort-Object name, version, publisher, source, registry_key |
  Group-Object name, version, publisher |
  ForEach-Object {
    $_.Group | Select-Object -First 1
  } |
  Sort-Object name, version, publisher

# Konverter til JSON og gem direkte på skrivebordet med UTF8-kodning
$deduped |
  ConvertTo-Json -Depth 4 |
  Out-File $outputPath -Encoding UTF8

Write-Host "Wrote $($deduped.Count) programs to $outputPath"
