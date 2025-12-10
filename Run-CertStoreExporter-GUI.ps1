<#
.SYNOPSIS
    GUI tool to export a certificate (User/Computer Personal) to PFX
    and generate Linux-ready PEM files using OpenSSL.

.DESCRIPTION
    This WinForms tool:
      1) Enumerates certificates from:
         - Cert:\CurrentUser\My
         - Cert:\LocalMachine\My
      2) Lets you select a certificate with a private key
      3) Exports it to PFX using Export-PfxCertificate
      4) Uses OpenSSL to create:
         - cert.pem
         - privkey.pem
         - chain.pem
         - fullchain.pem

.VERSION
    1.0

.AUTHOR
    Peter Schmidt (msdigest.net)

.LAST UPDATED
    2025-12-10

.NOTES
    - Output defaults to current folder (.\)
    - The PFX password is used to feed OpenSSL. The script converts it to plain text
      in memory only to pass -passin. If you prefer interactive prompts, you can
      toggle $UsePassIn = $false in the conversion function.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Helper: Find OpenSSL
# ----------------------------
function Get-OpenSslPath {
    # Try PATH
    $cmd = Get-Command openssl.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    # Common Git for Windows locations
    $candidates = @(
        "$env:ProgramFiles\Git\usr\bin\openssl.exe",
        "$env:ProgramFiles(x86)\Git\usr\bin\openssl.exe",
        "$env:LOCALAPPDATA\Programs\Git\usr\bin\openssl.exe"
    )

    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }

    return $null
}

# ----------------------------
# Helper: Load certs
# ----------------------------
function Get-PersonalStoreCerts {
    $stores = @(
        @{ Name = "CurrentUser"; Path = "Cert:\CurrentUser\My" },
        @{ Name = "LocalMachine"; Path = "Cert:\LocalMachine\My" }
    )

    $result = New-Object System.Collections.Generic.List[object]

    foreach ($s in $stores) {
        try {
            $certs = Get-ChildItem -Path $s.Path -ErrorAction Stop
            foreach ($c in $certs) {
                $result.Add([pscustomobject]@{
                    Store          = $s.Name
                    Path           = "$($s.Path)\$($c.Thumbprint)"
                    Subject        = $c.Subject
                    FriendlyName   = $c.FriendlyName
                    Thumbprint     = $c.Thumbprint
                    NotAfter       = $c.NotAfter
                    HasPrivateKey  = $c.HasPrivateKey
                })
            }
        } catch {
            # Ignore store read errors for UI stability
        }
    }

    return $result
}

# ----------------------------
# Helper: Export PFX
# ----------------------------
function Export-SelectedCertToPfx {
    param(
        [Parameter(Mandatory)]
        [string]$CertPath,

        [Parameter(Mandatory)]
        [string]$PfxFilePath,

        [Parameter(Mandatory)]
        [securestring]$Password
    )

    # Build chain into PFX when possible (helps chain.pem extraction)
    Export-PfxCertificate -Cert $CertPath `
        -FilePath $PfxFilePath `
        -Password $Password `
        -ChainOption BuildChain `
        -Force | Out-Null
}

# ----------------------------
# Helper: Convert PFX -> PEM set
# ----------------------------
function Convert-PfxToLinuxPem {
    param(
        [Parameter(Mandatory)]
        [string]$OpenSsl,

        [Parameter(Mandatory)]
        [string]$PfxFile,

        [Parameter(Mandatory)]
        [securestring]$Password,

        [Parameter(Mandatory)]
        [string]$OutputDir
    )

    if (!(Test-Path $OpenSsl)) {
        throw "OpenSSL not found at: $OpenSsl"
    }
    if (!(Test-Path $PfxFile)) {
        throw "PFX file not found: $PfxFile"
    }
    if (!(Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $base = [IO.Path]::GetFileNameWithoutExtension($PfxFile)

    $certPem      = Join-Path $OutputDir "$base-cert.pem"
    $keyPem       = Join-Path $OutputDir "$base-privkey.pem"
    $chainPem     = Join-Path $OutputDir "$base-chain.pem"
    $fullchainPem = Join-Path $OutputDir "$base-fullchain.pem"

    # SECURITY NOTE:
    # OpenSSL supports -passin. To avoid extra prompts in a GUI,
    # we convert the SecureString to plain text in memory.
    $UsePassIn = $true

    $plain = ""
    if ($UsePassIn) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }

    $passArgs = @()
    if ($UsePassIn -and $plain) {
        $passArgs = @("-passin", "pass:$plain")
    }

    # 1) cert.pem (leaf)
    & $OpenSsl pkcs12 -in $PfxFile -clcerts -nokeys -out $certPem @passArgs 2>&1 | Out-Null

    # 2) privkey.pem
    & $OpenSsl pkcs12 -in $PfxFile -nocerts -nodes -out $keyPem @passArgs 2>&1 | Out-Null

    # 3) chain.pem (intermediates)
    & $OpenSsl pkcs12 -in $PfxFile -cacerts -nokeys -out $chainPem @passArgs 2>&1 | Out-Null

    # 4) fullchain.pem (cert + chain)
    $certContent  = Get-Content -Path $certPem  -ErrorAction SilentlyContinue
    $chainContent = Get-Content -Path $chainPem -ErrorAction SilentlyContinue
    @($certContent + $chainContent) | Set-Content -Path $fullchainPem -Encoding ascii

    # Return paths for UI display
    return [pscustomobject]@{
        CertPem      = $certPem
        PrivateKey   = $keyPem
        ChainPem     = $chainPem
        FullchainPem = $fullchainPem
        OutputDir    = $OutputDir
    }
}

# ----------------------------
# UI Construction
# ----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export Certificate to Linux PEM Set"
$form.Size = New-Object System.Drawing.Size(950, 620)
$form.StartPosition = "CenterScreen"

# Store label
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Select a certificate from Personal stores (User/Computer). Only certs with private keys can be exported to PFX."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$form.Controls.Add($lblInfo)

# ListView
$list = New-Object System.Windows.Forms.ListView
$list.View = "Details"
$list.FullRowSelect = $true
$list.GridLines = $true
$list.Location = New-Object System.Drawing.Point(12, 40)
$list.Size = New-Object System.Drawing.Size(910, 380)

@("Store","FriendlyName","Subject","Thumbprint","Expires","HasPrivateKey") | ForEach-Object {
    $col = New-Object System.Windows.Forms.ColumnHeader
    $col.Text = $_
    $col.Width = switch ($_) {
        "Store" { 110 }
        "FriendlyName" { 170 }
        "Subject" { 230 }
        "Thumbprint" { 240 }
        "Expires" { 110 }
        "HasPrivateKey" { 90 }
        default { 120 }
    }
    $list.Columns.Add($col) | Out-Null
}

$form.Controls.Add($list)

# OpenSSL path box
$lblOpenSsl = New-Object System.Windows.Forms.Label
$lblOpenSsl.Text = "OpenSSL path:"
$lblOpenSsl.AutoSize = $true
$lblOpenSsl.Location = New-Object System.Drawing.Point(12, 435)
$form.Controls.Add($lblOpenSsl)

$txtOpenSsl = New-Object System.Windows.Forms.TextBox
$txtOpenSsl.Location = New-Object System.Drawing.Point(110, 432)
$txtOpenSsl.Size = New-Object System.Drawing.Size(650, 25)
$form.Controls.Add($txtOpenSsl)

$btnBrowseOpenSsl = New-Object System.Windows.Forms.Button
$btnBrowseOpenSsl.Text = "Browse..."
$btnBrowseOpenSsl.Location = New-Object System.Drawing.Point(770, 430)
$btnBrowseOpenSsl.Size = New-Object System.Drawing.Size(75, 28)
$form.Controls.Add($btnBrowseOpenSsl)

# Buttons
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(12, 475)
$btnRefresh.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($btnRefresh)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Selected → PFX → Linux PEMs"
$btnExport.Location = New-Object System.Drawing.Point(140, 475)
$btnExport.Size = New-Object System.Drawing.Size(280, 35)
$form.Controls.Add($btnExport)

# Status box
$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Multiline = $true
$txtStatus.ReadOnly = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.Location = New-Object System.Drawing.Point(12, 520)
$txtStatus.Size = New-Object System.Drawing.Size(910, 60)
$form.Controls.Add($txtStatus)

function Write-Status($msg) {
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $txtStatus.AppendText("[$timestamp] $msg`r`n")
}

# Populate OpenSSL path on load
$autoOpenSsl = Get-OpenSslPath
if ($autoOpenSsl) { $txtOpenSsl.Text = $autoOpenSsl }

# Load certs into list
function Load-CertList {
    $list.Items.Clear()
    $certs = Get-PersonalStoreCerts | Sort-Object Store, Subject

    foreach ($c in $certs) {
        $item = New-Object System.Windows.Forms.ListViewItem($c.Store)
        $item.SubItems.Add($c.FriendlyName) | Out-Null
        $item.SubItems.Add($c.Subject) | Out-Null
        $item.SubItems.Add($c.Thumbprint) | Out-Null
        $item.SubItems.Add(($c.NotAfter.ToString("yyyy-MM-dd"))) | Out-Null
        $item.SubItems.Add([string]$c.HasPrivateKey) | Out-Null

        # stash object for later
        $item.Tag = $c
        $list.Items.Add($item) | Out-Null
    }

    Write-Status "Loaded $($certs.Count) certificates from User/Computer Personal stores."
}

# Browse for OpenSSL
$btnBrowseOpenSsl.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "OpenSSL (openssl.exe)|openssl.exe|All files (*.*)|*.*"
    $dlg.InitialDirectory = (Get-Location).Path
    if ($dlg.ShowDialog() -eq "OK") {
        $txtOpenSsl.Text = $dlg.FileName
        Write-Status "OpenSSL path set to: $($dlg.FileName)"
    }
})

# Refresh list
$btnRefresh.Add_Click({
    Load-CertList
})

# Export workflow
$btnExport.Add_Click({
    try {
        if ($list.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select a certificate first.","No selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        $selected = $list.SelectedItems[0].Tag

        if (-not $selected.HasPrivateKey) {
            [System.Windows.Forms.MessageBox]::Show("The selected certificate does not have a private key. Cannot export to PFX.",
                "No private key",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $openSslPath = $txtOpenSsl.Text
        if ([string]::IsNullOrWhiteSpace($openSslPath) -or -not (Test-Path $openSslPath)) {
            [System.Windows.Forms.MessageBox]::Show("OpenSSL.exe not found. Please set the OpenSSL path.",
                "OpenSSL missing",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        # Ask for PFX save location
        $save = New-Object System.Windows.Forms.SaveFileDialog
        $save.Filter = "PFX files (*.pfx)|*.pfx|All files (*.*)|*.*"
        $save.InitialDirectory = (Get-Location).Path
        $safeName = ($selected.Subject -replace '[^\w\.-]+','_')
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = $selected.Thumbprint }
        $save.FileName = "$safeName.pfx"

        if ($save.ShowDialog() -ne "OK") { return }

        $pfxPath = $save.FileName

        # Prompt for PFX password
        $pwd1 = Read-Host "Enter PFX password for export (will also be used for OpenSSL)" -AsSecureString

        if (-not $pwd1) {
            [System.Windows.Forms.MessageBox]::Show("Password was empty. Aborting.",
                "Password required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        Write-Status "Exporting certificate from $($selected.Store) store..."
        Export-SelectedCertToPfx -CertPath $selected.Path -PfxFilePath $pfxPath -Password $pwd1
        Write-Status "PFX exported to: $pfxPath"

        # Choose output folder for PEMs (default = PFX folder)
        $outputDir = Split-Path $pfxPath -Parent
        if ([string]::IsNullOrWhiteSpace($outputDir)) {
            $outputDir = (Get-Location).Path
        }

        Write-Status "Generating Linux PEM files..."
        $result = Convert-PfxToLinuxPem -OpenSsl $openSslPath -PfxFile $pfxPath -Password $pwd1 -OutputDir $outputDir

        Write-Status "Created:"
        Write-Status "  cert.pem:      $($result.CertPem)"
        Write-Status "  privkey.pem:   $($result.PrivateKey)"
        Write-Status "  chain.pem:     $($result.ChainPem)"
        Write-Status "  fullchain.pem: $($result.FullchainPem)"

        [System.Windows.Forms.MessageBox]::Show(
            "Export complete.`r`n`r`ncert.pem`r`nprivkey.pem`r`nchain.pem`r`nfullchain.pem`r`n`r`nFolder:`r`n$($result.OutputDir)",
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

    } catch {
        Write-Status "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

# Initial load
Load-CertList

# Show UI
[void]$form.ShowDialog()
