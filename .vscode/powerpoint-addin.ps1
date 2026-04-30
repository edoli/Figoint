[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("build", "prepare")]
    [string]$Action,

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$workspaceRoot = Split-Path -Path $PSScriptRoot -Parent
$solutionPath = Join-Path $workspaceRoot "EdoliAddIn.sln"
$projectRoot = Join-Path $workspaceRoot "EdoliAddIn"
$registryPath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\EdoliAddIn"

function Get-MSBuildPath {
    $vswherePath = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswherePath)) {
        throw "vswhere.exe not found. Install Visual Studio 2022 or Build Tools with MSBuild."
    }

    $msbuildPath = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" |
        Select-Object -First 1

    if (-not $msbuildPath) {
        throw "MSBuild.exe not found. Install the MSBuild component."
    }

    return $msbuildPath
}

function Get-PowerPointPath {
    $installRoots = @(
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\PowerPoint\InstallRoot",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\PowerPoint\InstallRoot"
    )

    foreach ($installRoot in $installRoots) {
        if (-not (Test-Path $installRoot)) {
            continue
        }

        $officeRoot = Get-ItemPropertyValue -Path $installRoot -Name "Path" -ErrorAction SilentlyContinue
        if (-not $officeRoot) {
            continue
        }

        $powerPointPath = Join-Path $officeRoot "POWERPNT.EXE"
        if (Test-Path $powerPointPath) {
            return $powerPointPath
        }
    }

    throw "POWERPNT.EXE not found. Install Microsoft PowerPoint and update launch.json if it lives elsewhere."
}

function Assert-PowerPointClosed {
    $runningProcess = Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($runningProcess) {
        throw "PowerPoint is already running. Close all PowerPoint windows before starting this VS Code launch profile."
    }
}

function Invoke-Build {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BuildConfiguration
    )

    $msbuildPath = Get-MSBuildPath

    & $msbuildPath $solutionPath /p:Configuration=$BuildConfiguration /p:Platform="Any CPU" /m /nologo /verbosity:minimal /t:Build
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed for configuration '$BuildConfiguration'."
    }
}

function Set-ManifestRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BuildConfiguration
    )

    $manifestPath = Join-Path $projectRoot "bin\$BuildConfiguration\EdoliAddIn.vsto"
    if (-not (Test-Path $manifestPath)) {
        throw "Manifest not found at '$manifestPath'."
    }

    $manifestUri = "file:///$($manifestPath -replace '\\', '/')|vstolocal"

    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    New-ItemProperty -Path $registryPath -Name "FriendlyName" -PropertyType String -Value "EdoliAddIn" -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "Description" -PropertyType String -Value "EdoliAddIn" -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "Manifest" -PropertyType String -Value $manifestUri -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name "LoadBehavior" -PropertyType DWord -Value 3 -Force | Out-Null

    Write-Host "Registered manifest: $manifestUri"
}

switch ($Action) {
    "build" {
        Invoke-Build -BuildConfiguration $Configuration
    }
    "prepare" {
        Assert-PowerPointClosed
        Invoke-Build -BuildConfiguration $Configuration
        Set-ManifestRegistration -BuildConfiguration $Configuration
        $powerPointPath = Get-PowerPointPath
        Write-Host "PowerPoint path: $powerPointPath"
    }
}
