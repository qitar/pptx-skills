<#
.SYNOPSIS
    Bootstrap a virtualenv and run a pptx-compare strategy.

.DESCRIPTION
    PowerShell equivalent of run.sh, for Windows users without bash.
    Creates .venv next to this script on first run and installs
    requirements.txt into it, then dispatches to the chosen strategy.
    The strategy's JSON output is written to stdout unmodified.

.EXAMPLE
    .\run.ps1 lite before.pptx after.pptx

.EXAMPLE
    .\run.ps1 full before.pptx after.pptx --threshold 0.6
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('lite', 'full')]
    [string]$Strategy,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Rest = @()
)

# Deliberately NOT 'Stop': in Windows PowerShell 5.1 any stderr output from a
# native executable is wrapped in a NativeCommandError, which under 'Stop'
# aborts the script even when the command succeeded. pip and the Microsoft
# Store python stub both write to stderr routinely, so exit codes are checked
# explicitly instead.
$ErrorActionPreference = 'Continue'

$Here = $PSScriptRoot
$Venv = Join-Path $Here '.venv'

function Test-PythonCandidate {
    <#
        Return $true when invoking `$Exe $PreArgs` yields Python >= 3.10.
        Output is discarded: the Store stub prints a banner and exits 9009.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Exe,
        [string[]]$PreArgs = @()
    )
    $probe = 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'
    & $Exe @($PreArgs + @('-c', $probe)) 2>&1 | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Find-Python {
    foreach ($name in @('python3.13', 'python3.12', 'python3.11', 'python3.10', 'python3', 'python')) {
        if (-not (Get-Command $name -ErrorAction SilentlyContinue)) { continue }
        if (Test-PythonCandidate -Exe $name) {
            return @{ Cmd = $name; Args = @() }
        }
    }
    # Windows py launcher, if the plain names were unusable.
    if (Get-Command 'py' -ErrorAction SilentlyContinue) {
        if (Test-PythonCandidate -Exe 'py' -PreArgs @('-3')) {
            return @{ Cmd = 'py'; Args = @('-3') }
        }
    }
    return $null
}

function Find-VenvPython {
    # Windows layout first, POSIX layout for PowerShell on Linux/macOS.
    foreach ($rel in @('Scripts/python.exe', 'bin/python')) {
        $path = Join-Path $Venv $rel
        if (Test-Path $path) { return $path }
    }
    return $null
}

$VenvPython = Find-VenvPython
if ($null -eq $VenvPython) {
    $python = Find-Python
    if ($null -eq $python) {
        Write-Error 'Python >= 3.10 is required but was not found on PATH.'
        exit 1
    }
    & $python.Cmd @($python.Args + @('-m', 'venv', '--clear', $Venv))
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to create a virtualenv at $Venv"
        exit 1
    }
    $VenvPython = Find-VenvPython
}
if ($null -eq $VenvPython) {
    Write-Error "Could not locate the virtualenv interpreter under $Venv"
    exit 1
}

& $VenvPython -m pip install -r (Join-Path $Here 'requirements.txt') -q
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Failed to install dependencies from requirements.txt'
    exit 1
}

$MainScript = if ($Strategy -eq 'lite') {
    Join-Path $Here 'slide-compare-lite/scripts/main.py'
} else {
    Join-Path $Here 'slide-compare/scripts/main.py'
}

$env:PYTHONPATH = $Here
& $VenvPython $MainScript @Rest
exit $LASTEXITCODE
