#InstallWordTools.ps1
#This is in Barracuda RMM. It deploys Rick Tools MS Office ribbon.
#requires -Version 5.1
#requires -RunAsAdministrator

[CmdletBinding()]
param()

# Randomly stagger RMM downloads across devices by 0–60 seconds.
$staggerSeconds = Get-Random -Minimum 0 -Maximum 61
Write-Output "Waiting $staggerSeconds seconds"
Start-Sleep -Seconds $staggerSeconds

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$expectedCertificateThumbprint = '3E80DB47C38D0A620ECD8B5B596628BFA6362C3C'

$wordTemplateUrl = 'https://rickpubfiles.blob.core.windows.net/files/RICKWordTools.dotm'
$excelAddInUrl = 'https://rickpubfiles.blob.core.windows.net/files/RickExcelTools.xlam'
$certificateUrl = 'https://rickpubfiles.blob.core.windows.net/files/RICK-VBA-Code-Signing.cer'
$rulesUrl = 'https://rickpubfiles.blob.core.windows.net/files/rules.xml'
$roleMembersUrl = 'https://rickpubfiles.blob.core.windows.net/files/rolemembers.xml?sp=r&st=2026-09-23T19:17:43Z&se=2056-09-24T03:32:43Z&spr=https&sv=2026-02-06&sr=b&sig=i6N3ctruruA96C6mW5otwGyNctS%2FST4Hxgwu1Klg09Q%3D'

$rulesDirectory = Join-Path $env:ProgramData 'RICK Engineering\WordTools'
$installedRulesPath = Join-Path $rulesDirectory 'rules.xml'
$installedRoleMembersPath = Join-Path $rulesDirectory 'rolemembers.xml'
$logDirectory = Join-Path $rulesDirectory 'Logs'
$monthlyLogPath = Join-Path $logDirectory (
    'Install-{0}.log' -f (Get-Date).ToString('yyyy-MM')
)

$expiredLogPath = Join-Path $logDirectory (
    'Install-{0}.log' -f (Get-Date).AddMonths(-2).ToString('yyyy-MM')
)

New-Item `
    -Path $logDirectory `
    -ItemType Directory `
    -Force |
    Out-Null

# Retain the current and previous month's logs.
if (Test-Path $expiredLogPath){
    Remove-Item `
    -LiteralPath $expiredLogPath `
    -Force `
    -ErrorAction SilentlyContinue
}

$stagingDirectory = Join-Path $env:TEMP (
    'RICKOfficeTools-' + [guid]::NewGuid().ToString('N')
)

$stagedWordTemplatePath = Join-Path $stagingDirectory 'RICKWordTools.dotm'
$stagedExcelAddInPath = Join-Path $stagingDirectory 'RickExcelTools.xlam'
$stagedCertificatePath = Join-Path $stagingDirectory 'RICK-VBA-Code-Signing.cer'
$stagedRulesPath = Join-Path $stagingDirectory 'rules.xml'
$stagedRoleMembersPath = Join-Path $stagingDirectory 'rolemembers.xml'

$deploymentStatus = [ordered]@{
    Certificate     = 'Not attempted'
    Rules           = 'Not attempted'
    RoleMemberships = 'Not attempted'
    WordAddIn       = 'Not attempted'
    ExcelAddIn      = 'Not attempted'
}

$failureCount = 0
$rulesVersion = $null
$installedWordTemplatePath = $null
$installedExcelAddInPath = $null

function Write-DeploymentError {
    param(
        [Parameter(Mandatory)]
        [string]$Component,

        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $script:failureCount++
    $script:deploymentStatus[$Component] = 'Failed: ' + $ErrorRecord.Exception.Message

    $errorEntry = @(
        '------------------------------------------------------------'
        "Timestamp: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        "Computer:  $env:COMPUTERNAME"
        "User:      $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "Component: $Component"
        "Error:     $($ErrorRecord.Exception.Message)"
        "Location:  $($ErrorRecord.InvocationInfo.PositionMessage)"
    ) -join [Environment]::NewLine

    try {
        Add-Content `
            -LiteralPath $script:monthlyLogPath `
            -Value $errorEntry `
            -Encoding UTF8
    }
    catch {
    }

    Write-Output "$Component failed: $($ErrorRecord.Exception.Message)"
}

function Invoke-RickDownload {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    Invoke-WebRequest `
        -Uri $Uri `
        -OutFile $Destination `
        -UseBasicParsing

    if (
        -not (Test-Path -LiteralPath $Destination -PathType Leaf) -or
        (Get-Item -LiteralPath $Destination).Length -eq 0
    ) {
        throw "Downloaded file is missing or empty: $Destination"
    }
}

function Get-OfficeExecutablePath {
    param(
        [Parameter(Mandatory)]
        [string]$ExecutableName
    )

    $candidates = @()

    foreach ($registryPath in @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExecutableName"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExecutableName"
    )) {
        if (Test-Path -LiteralPath $registryPath) {
            $registryKey = Get-Item -LiteralPath $registryPath
            $registeredPath = $registryKey.GetValue('')

            if ($registeredPath) {
                $candidates += $registeredPath
            }
        }
    }

    if ($env:ProgramFiles) {
        $candidates += Join-Path $env:ProgramFiles (
            "Microsoft Office\root\Office16\$ExecutableName"
        )
    }

    $programFilesX86 = ${env:ProgramFiles(x86)}

    if ($programFilesX86) {
        $candidates += Join-Path $programFilesX86 (
            "Microsoft Office\root\Office16\$ExecutableName"
        )
    }

    $executablePath = $candidates |
        Where-Object {
            $_ -and (Test-Path -LiteralPath $_ -PathType Leaf)
        } |
        Select-Object -First 1

    if (-not $executablePath) {
        throw "$ExecutableName could not be located on this computer."
    }

    return $executablePath
}

try {
    New-Item `
        -Path $stagingDirectory `
        -ItemType Directory `
        -Force |
        Out-Null
}
catch {
    Write-Output "RICK Office Tools could not initialize staging: $($_.Exception.Message)"
    Start-Sleep -Seconds 30
    exit 0
}

# Force TLS 1.2 for Windows PowerShell 5.1.
[Net.ServicePointManager]::SecurityProtocol = `
    [Net.ServicePointManager]::SecurityProtocol -bor `
    [Net.SecurityProtocolType]::Tls12

Write-Output 'Starting RICK Office Tools deployment.'

#Install the cert if it doesn't exist already
try {
    Write-Output 'Downloading and validating the signing certificate.'

    Invoke-RickDownload `
        -Uri $certificateUrl `
        -Destination $stagedCertificatePath

    $downloadedCertificate = New-Object `
        System.Security.Cryptography.X509Certificates.X509Certificate2(
            $stagedCertificatePath
        )

    $actualThumbprint = (
        $downloadedCertificate.Thumbprint -replace '\s', ''
    ).ToUpperInvariant()

    if ($actualThumbprint -ne $expectedCertificateThumbprint) {
        throw (
            'Certificate thumbprint mismatch. ' +
            "Expected $expectedCertificateThumbprint but received " +
            "$actualThumbprint."
        )
    }

    foreach ($storeName in @('Root', 'TrustedPublisher')) {

        $certificateStore = New-Object `
            System.Security.Cryptography.X509Certificates.X509Store(
                $storeName,
                [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
            )

        try {
            $certificateStore.Open(
                [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite
            )

            $existingCertificate = $certificateStore.Certificates |
                Where-Object {
                    ($_.Thumbprint -replace '\s', '').ToUpperInvariant() -eq
                    $expectedCertificateThumbprint
                } |
                Select-Object -First 1

            if ($existingCertificate) {
                Write-Output "Certificate already exists in LocalMachine\$storeName."
            }
            else {
                $certificateStore.Add($downloadedCertificate)
                Write-Output "Certificate installed in LocalMachine\$storeName."
            }
        }
        finally {
            $certificateStore.Close()
        }
    }

    $deploymentStatus.Certificate = "Installed/verified: $actualThumbprint"
}
catch {
    Write-DeploymentError -Component 'Certificate' -ErrorRecord $_
}

#Download the rules.xml
try {
    Write-Output 'Downloading and validating rules.xml.'

    Invoke-RickDownload `
        -Uri $rulesUrl `
        -Destination $stagedRulesPath

    $rulesXml = New-Object System.Xml.XmlDocument
    $rulesXml.PreserveWhitespace = $true
    $rulesXml.Load($stagedRulesPath)

    if (
        $null -eq $rulesXml.DocumentElement -or
        $rulesXml.DocumentElement.Name -ne 'rickWordRules'
    ) {
        throw 'The downloaded rules file does not have the expected rickWordRules root element.'
    }

    $ruleCount = @(
        $rulesXml.SelectNodes('/rickWordRules/rules/rule')
    ).Count

    if ($ruleCount -lt 1) {
        throw 'The downloaded rules file contains no rules.'
    }

    $rulesVersion = $rulesXml.DocumentElement.GetAttribute('version')

    New-Item `
        -Path $rulesDirectory `
        -ItemType Directory `
        -Force |
        Out-Null

    Copy-Item `
        -LiteralPath $stagedRulesPath `
        -Destination $installedRulesPath `
        -Force

    $deploymentStatus.Rules = (
        "Installed: version $rulesVersion; rules $ruleCount"
    )
}
catch {
    Write-DeploymentError -Component 'Rules' -ErrorRecord $_
}

#Download the rolemembers.xml
try {
    Write-Output 'Downloading and validating rolemembers.xml.'

    Invoke-RickDownload `
        -Uri $roleMembersUrl `
        -Destination $stagedRoleMembersPath

    $roleMembersXml = New-Object System.Xml.XmlDocument
    $roleMembersXml.Load($stagedRoleMembersPath)

    if (
        $null -eq $roleMembersXml.DocumentElement -or
        $roleMembersXml.DocumentElement.Name -ne 'rickOfficeRoleMembers'
    ) {
        throw 'The downloaded role membership file does not have the expected rickOfficeRoleMembers root element.'
    }

    New-Item `
        -Path $rulesDirectory `
        -ItemType Directory `
        -Force |
        Out-Null

    Copy-Item `
        -LiteralPath $stagedRoleMembersPath `
        -Destination $installedRoleMembersPath `
        -Force

    $deploymentStatus.RoleMemberships = "Installed: $installedRoleMembersPath"
}
catch {
    Write-DeploymentError -Component 'RoleMemberships' -ErrorRecord $_
}

#Deploy the Word file
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    $deploymentStatus.WordAddIn = 'Skipped: Microsoft Word is running'
    Write-Output 'Word add-in skipped because Microsoft Word is running.'
}
else {
    try {
        Write-Output 'Downloading and installing the Word add-in.'

        Invoke-RickDownload `
            -Uri $wordTemplateUrl `
            -Destination $stagedWordTemplatePath

        $wordExecutable = Get-OfficeExecutablePath -ExecutableName 'WINWORD.EXE'
        $wordDirectory = Split-Path -Path $wordExecutable -Parent
        $wordStartupDirectory = Join-Path $wordDirectory 'STARTUP'

        $installedWordTemplatePath = Join-Path `
            $wordStartupDirectory `
            'RICKWordTools.dotm'

        New-Item `
            -Path $wordStartupDirectory `
            -ItemType Directory `
            -Force |
            Out-Null

        Copy-Item `
            -LiteralPath $stagedWordTemplatePath `
            -Destination $installedWordTemplatePath `
            -Force

        $deploymentStatus.WordAddIn = "Installed: $installedWordTemplatePath"
    }
    catch {
        Write-DeploymentError -Component 'WordAddIn' -ErrorRecord $_
    }
}

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    $deploymentStatus.ExcelAddIn = 'Skipped: Microsoft Excel is running'
    Write-Output 'Excel add-in skipped because Microsoft Excel is running.'
}
else {
    try {
        Write-Output 'Downloading and installing the Excel add-in.'

        Invoke-RickDownload `
            -Uri $excelAddInUrl `
            -Destination $stagedExcelAddInPath

        $excelExecutable = Get-OfficeExecutablePath -ExecutableName 'EXCEL.EXE'
        $excelDirectory = Split-Path -Path $excelExecutable -Parent
        $excelStartupDirectory = Join-Path $excelDirectory 'XLSTART'

        $installedExcelAddInPath = Join-Path `
            $excelStartupDirectory `
            'RickExcelTools.xlam'

        New-Item `
            -Path $excelStartupDirectory `
            -ItemType Directory `
            -Force |
            Out-Null

        Copy-Item `
            -LiteralPath $stagedExcelAddInPath `
            -Destination $installedExcelAddInPath `
            -Force

        $deploymentStatus.ExcelAddIn = "Installed: $installedExcelAddInPath"
    }
    catch {
        Write-DeploymentError -Component 'ExcelAddIn' -ErrorRecord $_
    }
}

#Now clean up
try {
    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item `
            -LiteralPath $stagingDirectory `
            -Recurse `
            -Force `
            -ErrorAction Stop
    }
}
catch {
    Write-Output "Staging cleanup failed: $($_.Exception.Message)"
}

Write-Output ''
Write-Output 'RICK Office Tools deployment summary:'

foreach ($component in $deploymentStatus.Keys) {
    Write-Output ('{0}: {1}' -f $component, $deploymentStatus[$component])
}

if ($failureCount -gt 0) {
    Write-Output ''
    Write-Output "$failureCount component(s) failed. Other components were still attempted."
    Write-Output "Failure details: $monthlyLogPath"
}
else {
    Write-Output ''
    Write-Output 'All available components completed or were safely skipped.'
}

Start-Sleep -Seconds 30
exit 0











<#
try {

    Write-Output 'Starting RICK Word Tools installation.'

    # Avoid replacing a Word add-in while Word has it loaded.
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        throw 'Microsoft Word is currently running. Close Word and run the installation again.'
    }

    # Force TLS 1.2 for Windows PowerShell 5.1.
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor `
        [Net.SecurityProtocolType]::Tls12

    New-Item `
        -Path $stagingDirectory `
        -ItemType Directory `
        -Force |
        Out-Null

    Write-Output 'Downloading deployment files.'

    Invoke-WebRequest `
        -Uri $templateUrl `
        -OutFile $stagedTemplatePath `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri $certificateUrl `
        -OutFile $stagedCertificatePath `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri $rulesUrl `
        -OutFile $stagedRulesPath `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri $roleMembersUrl `
        -OutFile $stagedRoleMembersPath `
        -UseBasicParsing

    # Confirm all downloads produced nonempty files.
    foreach ($downloadedFile in @(
        $stagedTemplatePath
        $stagedCertificatePath
        $stagedRulesPath
        $stagedRoleMembersPath
    )) {
        if (
            -not (Test-Path -LiteralPath $downloadedFile -PathType Leaf) -or
            (Get-Item -LiteralPath $downloadedFile).Length -eq 0
        ) {
            throw "Downloaded file is missing or empty: $downloadedFile"
        }
    }

    Write-Output 'Validating signing certificate.'

    $downloadedCertificate = New-Object `
        System.Security.Cryptography.X509Certificates.X509Certificate2(
            $stagedCertificatePath
        )

    $actualThumbprint = (
        $downloadedCertificate.Thumbprint -replace '\s', ''
    ).ToUpperInvariant()

    if ($actualThumbprint -ne $expectedCertificateThumbprint) {
        throw (
            'Certificate thumbprint mismatch. ' +
            "Expected $expectedCertificateThumbprint but received " +
            "$actualThumbprint."
        )
    }

    Write-Output "Certificate verified: $actualThumbprint"

    Write-Output 'Validating rules XML.'

    $rulesXml = New-Object System.Xml.XmlDocument
    $rulesXml.PreserveWhitespace = $true
    $rulesXml.Load($stagedRulesPath)

    if (
        $null -eq $rulesXml.DocumentElement -or
        $rulesXml.DocumentElement.Name -ne 'rickWordRules'
    ) {
        throw 'The downloaded rules file does not have the expected rickWordRules root element.'
    }

    $ruleCount = @(
        $rulesXml.SelectNodes('/rickWordRules/rules/rule')
    ).Count

    if ($ruleCount -lt 1) {
        throw 'The downloaded rules file contains no rules.'
    }

    $rulesVersion = $rulesXml.DocumentElement.GetAttribute('version')

    Write-Output "Rules validated. Version: $rulesVersion; Rules: $ruleCount"

    Write-Output 'Locating Microsoft Word.'

    $wordCandidates = @()

    foreach ($registryPath in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )) {
        if (Test-Path -LiteralPath $registryPath) {
            $registryKey = Get-Item -LiteralPath $registryPath
            $registeredWordPath = $registryKey.GetValue('')

            if ($registeredWordPath) {
                $wordCandidates += $registeredWordPath
            }
        }
    }

    if ($env:ProgramFiles) {
        $wordCandidates += Join-Path $env:ProgramFiles `
            'Microsoft Office\root\Office16\WINWORD.EXE'
    }

    $programFilesX86 = ${env:ProgramFiles(x86)}

    if ($programFilesX86) {
        $wordCandidates += Join-Path $programFilesX86 `
            'Microsoft Office\root\Office16\WINWORD.EXE'
    }

    $wordExecutable = $wordCandidates |
        Where-Object {
            $_ -and (Test-Path -LiteralPath $_ -PathType Leaf)
        } |
        Select-Object -First 1

    if (-not $wordExecutable) {
        throw 'Microsoft Word could not be located on this computer.'
    }

    $wordDirectory = Split-Path -Path $wordExecutable -Parent
    $wordStartupDirectory = Join-Path $wordDirectory 'STARTUP'
    $installedTemplatePath = Join-Path `
        $wordStartupDirectory `
        'RICKWordTools.dotm'

    Write-Output "Word found: $wordExecutable"
    Write-Output "Word startup directory: $wordStartupDirectory"

    New-Item `
        -Path $rulesDirectory `
        -ItemType Directory `
        -Force |
        Out-Null

    New-Item `
        -Path $wordStartupDirectory `
        -ItemType Directory `
        -Force |
        Out-Null

    Write-Output 'Installing public certificate.'

    foreach ($storeName in @('Root', 'TrustedPublisher')) {

        $certificateStore = New-Object `
            System.Security.Cryptography.X509Certificates.X509Store(
                $storeName,
                [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
            )

        try {
            $certificateStore.Open(
                [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite
            )

            $existingCertificate = $certificateStore.Certificates |
                Where-Object {
                    ($_.Thumbprint -replace '\s', '').ToUpperInvariant() -eq
                    $expectedCertificateThumbprint
                } |
                Select-Object -First 1

            if ($existingCertificate) {
                Write-Output "Certificate already exists in LocalMachine\$storeName."
            }
            else {
                $certificateStore.Add($downloadedCertificate)
                Write-Output "Certificate installed in LocalMachine\$storeName."
            }
        }
        finally {
            $certificateStore.Close()
        }
    }

    Write-Output 'Installing rules.xml.'

    Copy-Item `
        -LiteralPath $stagedRulesPath `
        -Destination $installedRulesPath `
        -Force

    Write-Output 'Installing Word template.'

    Copy-Item `
        -LiteralPath $stagedTemplatePath `
        -Destination $installedTemplatePath `
        -Force

        Write-Output 'Validating role membership XML.'

    $roleMembersXml = New-Object System.Xml.XmlDocument
    $roleMembersXml.Load($stagedRoleMembersPath)

    if (
        $null -eq $roleMembersXml.DocumentElement -or
        $roleMembersXml.DocumentElement.Name -ne 'rickOfficeRoleMembers'
    ) {
        throw 'The downloaded role membership file does not have the expected rickOfficeRoleMembers root element.'
    }

    Write-Output 'Role membership XML validated.'

    Write-Output 'Installing rolemembers.xml.'

    Copy-Item `
        -LiteralPath $stagedRoleMembersPath `
        -Destination $installedRoleMembersPath `
        -Force
        
    Write-Output "Role memberships: $installedRoleMembersPath"

    Write-Output ''
    Write-Output 'RICK Word Tools installation completed successfully.'
    Write-Output "Template: $installedTemplatePath"
    Write-Output "Rules: $installedRulesPath"
    Write-Output "Rules version: $rulesVersion"
    Write-Output "Certificate: $expectedCertificateThumbprint"

}
catch {
    $errorEntry = @(
        '------------------------------------------------------------'
        "Timestamp: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        "Computer:  $env:COMPUTERNAME"
        "User:      $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "Error:     $($_.Exception.Message)"
        "Location:  $($_.InvocationInfo.PositionMessage)"
    ) -join [Environment]::NewLine

    try {
        Add-Content `
            -LiteralPath $monthlyLogPath `
            -Value $errorEntry `
            -Encoding UTF8
    }
    catch {
        # Do not generate another RMM error if local logging also fails.
    }

    Write-Output 'RICK Word Tools was not installed. Details were written to the local log.'
}
finally {
    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item `
            -LiteralPath $stagingDirectory `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }
}
Start-Sleep 30
exit 0
#>