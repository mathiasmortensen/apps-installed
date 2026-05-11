[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$Token,

    [Parameter(Mandatory=$false)]
    [string]$ApiUrl
)

$ErrorActionPreference = "Stop"

# Find brugerens faktiske Desktop-mappe
$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)

if ([string]::IsNullOrWhiteSpace($desktopPath)) {
  throw "Could not resolve Desktop path."
}

if (-not (Test-Path -LiteralPath $desktopPath)) {
  throw "Desktop path does not exist: $desktopPath"
}

$outputPath = Join-Path $desktopPath "apps-installeret.json"

Write-Host "User:         $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Host "Desktop path: $desktopPath"
Write-Host "Output path:  $outputPath"
Write-Host "Current dir:  $PWD"

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
  Get-ItemProperty -Path $source.Path -ErrorAction SilentlyContinue |
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

# Konverter til JSON.
# @($deduped) sikrer, at output behandles som en samling.
# Vi bruger ConvertTo-Json -Compress til API-kaldet, da en pæn formatering ikke er nødvendig for maskiner, 
# men vi gemmer stadig en pæn udgave på disken, hvis det ønskes.
$json = @($deduped) | ConvertTo-Json -Depth 4

# Gem direkte på skrivebordet med UTF8-kodning.
$json | Out-File -LiteralPath $outputPath -Encoding UTF8
Write-Host "Wrote $($deduped.Count) programs to $outputPath"

if (-not [string]::IsNullOrWhiteSpace($ApiUrl) -and -not [string]::IsNullOrWhiteSpace($Token)) {
    Write-Host "Poster data til API: $ApiUrl"
    
    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    try {
        $response = Invoke-RestMethod -Uri $ApiUrl -Method Post -Headers $headers -Body $json
        Write-Host "Upload fuldført med succes!" -ForegroundColor Green
    }
    catch {
        Write-Error "Kunne ikke uploade til API: $_"
    }
} else {
    Write-Host "ApiUrl og/eller Token mangler - data sendes ikke til API'et." -ForegroundColor Yellow
}

Write-Host "Venter 5 sekunder før output-filen slettes..."
Start-Sleep -Seconds 5
Remove-Item -Path $outputPath -Force
Write-Host "Output-filen '$outputPath' er blevet slettet."