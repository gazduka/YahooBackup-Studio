# Outlook classic COM importer. No changes to Yahoo/IMAP. Creates/uses a folder in a selected local PST.
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$PstPath,
    [Parameter(Mandatory=$true)][string]$FolderName,
    [Parameter(Mandatory=$true)][string]$Ledger,
    [Parameter(Mandatory=$true)][string]$StopFile,
    [string]$LegacyLedger = '',
    [int]$MaxItems = 10
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}
$Source = [IO.Path]::GetFullPath($Source)
$PstPath = [IO.Path]::GetFullPath($PstPath)
if (-not (Test-Path -LiteralPath $PstPath -PathType Leaf)) { throw 'Il file PST non esiste.' }
if ([IO.Path]::GetExtension($PstPath) -ine '.pst') { throw 'La destinazione deve essere un PST.' }
if (-not (Test-Path -LiteralPath $Source -PathType Container)) { throw 'Cartella MSG non trovata.' }
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
$found = $null
foreach ($store in $namespace.Stores) {
    try {
        if ($store.FilePath -and [string]::Equals([IO.Path]::GetFullPath($store.FilePath), $PstPath, [StringComparison]::OrdinalIgnoreCase)) {
            $found = $store
            break
        }
    } catch {}
}
if ($null -eq $found) { throw "Questo PST non e' aperto in Outlook: $PstPath. Aprilo da File > Apri file di dati di Outlook, quindi riprova." }
$root = $found.GetRootFolder()
$folder = $null
foreach ($sub in $root.Folders) {
    if ([string]::Equals($sub.Name, $FolderName, [StringComparison]::OrdinalIgnoreCase)) { $folder = $sub; break }
}
if ($null -eq $folder) { $folder = $root.Folders.Add($FolderName) }
Write-Output "PST|$PstPath"
Write-Output "CARTELLA|$($folder.FolderPath)"
$processed = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($log in @($Ledger, $LegacyLedger)) {
    if ($log -and (Test-Path -LiteralPath $log -PathType Leaf)) {
        foreach ($line in [IO.File]::ReadLines($log)) {
            $name = $line.Trim()
            if ($name -match '\.msg$') { [void]$processed.Add($name) }
        }
    }
}
# The destination-specific marker (where present) closes most of the crash-before-log gap.
# A marker may also be found on items imported in a previous run.
$existingMarkers = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
try {
    foreach ($item in $folder.Items) {
        try {
            $prop = $item.UserProperties.Find('YahooBackupStudioSource')
            if ($prop -and $prop.Value) { [void]$existingMarkers.Add([string]$prop.Value) }
        } catch {}
    }
} catch {
    Write-Output "WARN|Non e' stato possibile leggere tutti i marcatori esistenti: $($_.Exception.Message)"
}
$files = @(Get-ChildItem -LiteralPath $Source -Filter '*.msg' -File | Sort-Object Name)
$total = $files.Count
Write-Output "FILE|$total"
$ok=0; $skip=0; $errors=0; $index=0
$errLog = "$Ledger.errors.log"
foreach ($file in $files) {
    if ($MaxItems -gt 0 -and $ok -ge $MaxItems) { Write-Output 'STOP|Limite prova raggiunto'; break }
    if (Test-Path -LiteralPath $StopFile) { Write-Output 'STOP|Interruzione richiesta'; break }
    $index++
    $identity = $file.Name
    if ($processed.Contains($identity) -or $existingMarkers.Contains($identity)) {
        $skip++
        if (-not $processed.Contains($identity)) {
            [IO.File]::AppendAllText($Ledger, "$identity`r`n", [Text.UTF8Encoding]::new($false))
            [void]$processed.Add($identity)
        }
    } else {
        $mail = $null
        $moved = $null
        try {
            $mail = $namespace.OpenSharedItem($file.FullName)
            if ($null -eq $mail -or [string]$mail.MessageClass -ne 'IPM.Note') { throw 'Il MSG non e una email IPM.Note' }
            # Add a persistent marker to the email before moving. Outlook writes the PST, not us.
            try {
                $mark = $mail.UserProperties.Find('YahooBackupStudioSource')
                if ($null -eq $mark) { $mark = $mail.UserProperties.Add('YahooBackupStudioSource', 1, $true) }
                $mark.Value = $identity
                $mail.Save()
            } catch {
                Write-Output "WARN|Marcatore $identity non salvato: $($_.Exception.Message)"
            }
            $moved = $mail.Move($folder)
            if ($null -eq $moved -or [string]$moved.MessageClass -ne 'IPM.Note') { throw 'Spostamento non confermato come email' }
            [IO.File]::AppendAllText($Ledger, "$identity`r`n", [Text.UTF8Encoding]::new($false))
            [void]$processed.Add($identity)
            $ok++
        } catch {
            $errors++
            $detail = "${identity}: $($_.Exception.Message)"
            [IO.File]::AppendAllText($errLog, "$detail`r`n", [Text.UTF8Encoding]::new($false))
            Write-Output "ERRORE|$detail"
        } finally {
            if ($null -ne $mail) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail) } catch {} }
            if ($null -ne $moved) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($moved) } catch {} }
        }
    }
    if ($index % 25 -eq 0 -or $index -eq $total) {
        $pc = [math]::Round(100.0 * $index / [math]::Max(1, $total), 2)
        Write-Output "PROGRESS|$pc|$ok|$skip|$errors"
        Write-Output "[$index/$total] nuovi=$ok saltati=$skip errori=$errors"
    }
}
Write-Output "FINE|nuovi=$ok|saltati=$skip|errori=$errors|cartella=$($folder.FolderPath)"
if ($errors -gt 0) { exit 2 }
exit 0
