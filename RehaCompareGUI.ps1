Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Helper functions
# ----------------------------
function Select-FolderDialog([string]$InitialPath = "") {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath)) { $dlg.SelectedPath = $InitialPath }
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return $null
}

function Ensure-Folder([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
}

# .NET Framework-safe relative path
function Get-RelPath([string]$Root, [string]$FullName) {
    $root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\')
    $full = $FullName

    # If on different drive/root, fall back to full path
    if ($full.Length -lt 2 -or $root.Length -lt 2 -or $full.Substring(0,1) -ne $root.Substring(0,1)) {
        return $full
    }

    if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($root.Length).TrimStart('\')
    }

    # URI fallback
    try {
        $rootUri = New-Object System.Uri(($root.TrimEnd('\') + '\'))
        $fullUri = New-Object System.Uri($full)
        $relUri  = $rootUri.MakeRelativeUri($fullUri)
        return [System.Uri]::UnescapeDataString($relUri.ToString()).Replace('/', '\')
    } catch {
        return $full
    }
}

function Get-FileList([string]$Root, [bool]$Recurse) {
    $root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\')
    if ($Recurse) { Get-ChildItem -LiteralPath $root -Recurse -Force | Where-Object { -not $_.PSIsContainer } }
    else { Get-ChildItem -LiteralPath $root -Force | Where-Object { -not $_.PSIsContainer } }
}

function Safe-GetFileHash([string]$Path, [string]$Algorithm) {
    try { Get-FileHash -LiteralPath $Path -Algorithm $Algorithm }
    catch { $null }
}

function Write-Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logBox.AppendText("[$ts] $msg`r`n")
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ----------------------------
# GUI Layout
# ----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Compare (Name + Hash)"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"

$font = New-Object System.Drawing.Font("Segoe UI", 9)

# Folder A
$lblA = New-Object System.Windows.Forms.Label
$lblA.Text = "Folder A:"
$lblA.Location = New-Object System.Drawing.Point(12, 15)
$lblA.Size = New-Object System.Drawing.Size(70, 20)
$lblA.Font = $font
$form.Controls.Add($lblA)

$txtA = New-Object System.Windows.Forms.TextBox
$txtA.Location = New-Object System.Drawing.Point(85, 12)
$txtA.Size = New-Object System.Drawing.Size(680, 24)
$txtA.Font = $font
$txtA.Anchor = "Top,Left,Right"
$form.Controls.Add($txtA)

$btnA = New-Object System.Windows.Forms.Button
$btnA.Text = "Browse..."
$btnA.Location = New-Object System.Drawing.Point(775, 10)
$btnA.Size = New-Object System.Drawing.Size(95, 28)
$btnA.Font = $font
$form.Controls.Add($btnA)

# Folder B
$lblB = New-Object System.Windows.Forms.Label
$lblB.Text = "Folder B:"
$lblB.Location = New-Object System.Drawing.Point(12, 50)
$lblB.Size = New-Object System.Drawing.Size(70, 20)
$lblB.Font = $font
$form.Controls.Add($lblB)

$txtB = New-Object System.Windows.Forms.TextBox
$txtB.Location = New-Object System.Drawing.Point(85, 47)
$txtB.Size = New-Object System.Drawing.Size(680, 24)
$txtB.Font = $font
$txtB.Anchor = "Top,Left,Right"
$form.Controls.Add($txtB)

$btnB = New-Object System.Windows.Forms.Button
$btnB.Text = "Browse..."
$btnB.Location = New-Object System.Drawing.Point(775, 45)
$btnB.Size = New-Object System.Drawing.Size(95, 28)
$btnB.Font = $font
$form.Controls.Add($btnB)

# Output folder
$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Output:"
$lblOut.Location = New-Object System.Drawing.Point(12, 85)
$lblOut.Size = New-Object System.Drawing.Size(70, 20)
$lblOut.Font = $font
$form.Controls.Add($lblOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(85, 82)
$txtOut.Size = New-Object System.Drawing.Size(680, 24)
$txtOut.Font = $font
$txtOut.Anchor = "Top,Left,Right"
$form.Controls.Add($txtOut)

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = "Browse..."
$btnOut.Location = New-Object System.Drawing.Point(775, 80)
$btnOut.Size = New-Object System.Drawing.Size(95, 28)
$btnOut.Font = $font
$form.Controls.Add($btnOut)

# Case Number
$lblCase = New-Object System.Windows.Forms.Label
$lblCase.Text = "Case #:"
$lblCase.Location = New-Object System.Drawing.Point(12, 115)
$lblCase.Size = New-Object System.Drawing.Size(70, 20)
$lblCase.Font = $font
$form.Controls.Add($lblCase)

$txtCase = New-Object System.Windows.Forms.TextBox
$txtCase.Location = New-Object System.Drawing.Point(85, 112)
$txtCase.Size = New-Object System.Drawing.Size(250, 24)
$txtCase.Font = $font
$form.Controls.Add($txtCase)

# Operator
$lblOp = New-Object System.Windows.Forms.Label
$lblOp.Text = "Operator:"
$lblOp.Location = New-Object System.Drawing.Point(355, 115)
$lblOp.Size = New-Object System.Drawing.Size(70, 20)
$lblOp.Font = $font
$form.Controls.Add($lblOp)

$txtOp = New-Object System.Windows.Forms.TextBox
$txtOp.Location = New-Object System.Drawing.Point(430, 112)
$txtOp.Size = New-Object System.Drawing.Size(335, 24)
$txtOp.Font = $font
$form.Controls.Add($txtOp)

# Options group
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Options"
$grp.Location = New-Object System.Drawing.Point(12, 145)
$grp.Size = New-Object System.Drawing.Size(858, 140)
$grp.Font = $font
$grp.Anchor = "Top,Left,Right"
$form.Controls.Add($grp)

$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text = "Include subfolders (Recurse)"
$chkRecurse.Location = New-Object System.Drawing.Point(15, 25)
$chkRecurse.Size = New-Object System.Drawing.Size(250, 20)
$chkRecurse.Checked = $true
$chkRecurse.Font = $font
$grp.Controls.Add($chkRecurse)

$chkRelPath = New-Object System.Windows.Forms.CheckBox
$chkRelPath.Text = "Compare relative path (recommended)"
$chkRelPath.Location = New-Object System.Drawing.Point(15, 50)
$chkRelPath.Size = New-Object System.Drawing.Size(300, 20)
$chkRelPath.Checked = $true
$chkRelPath.Font = $font
$grp.Controls.Add($chkRelPath)

$chkNameOnly = New-Object System.Windows.Forms.CheckBox
$chkNameOnly.Text = "Compare filename-only (may have collisions)"
$chkNameOnly.Location = New-Object System.Drawing.Point(15, 75)
$chkNameOnly.Size = New-Object System.Drawing.Size(330, 20)
$chkNameOnly.Checked = $false
$chkNameOnly.Font = $font
$grp.Controls.Add($chkNameOnly)

$chkHash = New-Object System.Windows.Forms.CheckBox
$chkHash.Text = "Compute hashes and compare content"
$chkHash.Location = New-Object System.Drawing.Point(15, 100)
$chkHash.Size = New-Object System.Drawing.Size(280, 20)
$chkHash.Checked = $true
$chkHash.Font = $font
$grp.Controls.Add($chkHash)

$chkSameHashDiffPath = New-Object System.Windows.Forms.CheckBox
$chkSameHashDiffPath.Text = "Report: same hash but different path/name"
$chkSameHashDiffPath.Location = New-Object System.Drawing.Point(350, 100)
$chkSameHashDiffPath.Size = New-Object System.Drawing.Size(320, 20)
$chkSameHashDiffPath.Checked = $true
$chkSameHashDiffPath.Font = $font
$grp.Controls.Add($chkSameHashDiffPath)

$lblAlg = New-Object System.Windows.Forms.Label
$lblAlg.Text = "Hash algorithm:"
$lblAlg.Location = New-Object System.Drawing.Point(350, 25)
$lblAlg.Size = New-Object System.Drawing.Size(110, 20)
$lblAlg.Font = $font
$grp.Controls.Add($lblAlg)

$cmbAlg = New-Object System.Windows.Forms.ComboBox
$cmbAlg.Location = New-Object System.Drawing.Point(460, 22)
$cmbAlg.Size = New-Object System.Drawing.Size(130, 24)
$cmbAlg.DropDownStyle = "DropDownList"
[void]$cmbAlg.Items.AddRange(@("SHA256","SHA1","MD5"))
$cmbAlg.SelectedItem = "SHA256"
$cmbAlg.Font = $font
$grp.Controls.Add($cmbAlg)

# Run button + Progress
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Comparison"
$btnRun.Location = New-Object System.Drawing.Point(12, 305)
$btnRun.Size = New-Object System.Drawing.Size(160, 34)
$btnRun.Font = $font
$btnRun.Anchor = "Top,Left"
$form.Controls.Add($btnRun)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(185, 305)
$progress.Size = New-Object System.Drawing.Size(685, 24)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$progress.Anchor = "Top,Left,Right"
$form.Controls.Add($progress)

# Log window
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 350)
$logBox.Size = New-Object System.Drawing.Size(858, 300)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($logBox)

# Metadata
$AppName    = "RehaCompareGUIv2"
$AppVersion = "2.3"
$BuildDate  = "2026-02-11"

# ----------------------------
# Button handlers
# ----------------------------
$btnA.Add_Click({ $sel = Select-FolderDialog $txtA.Text; if ($sel) { $txtA.Text = $sel } })
$btnB.Add_Click({ $sel = Select-FolderDialog $txtB.Text; if ($sel) { $txtB.Text = $sel } })
$btnOut.Add_Click({ $sel = Select-FolderDialog $txtOut.Text; if ($sel) { $txtOut.Text = $sel } })

# Default output suggestion
$txtOut.Text = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

# ----------------------------
# Run logic (UI thread)
# ----------------------------
$btnRun.Add_Click({
    $progress.Value = 0
    $logBox.Clear()

    $CaseNumber   = $txtCase.Text.Trim()
    $Operator     = $txtOp.Text.Trim()
    $RunTimeLocal = Get-Date
    $RunTimeUtc   = (Get-Date).ToUniversalTime()

    $FolderA = $txtA.Text.Trim()
    $FolderB = $txtB.Text.Trim()
    $OutDir  = $txtOut.Text.Trim()

    $DoRecurse          = $chkRecurse.Checked
    $DoRelPath          = $chkRelPath.Checked
    $DoNameOnly         = $chkNameOnly.Checked
    $DoHash             = $chkHash.Checked
    $DoSameHashDiffPath = $chkSameHashDiffPath.Checked
    $Alg                = $cmbAlg.SelectedItem.ToString()

    if (-not (Test-Path -LiteralPath $FolderA)) { [System.Windows.Forms.MessageBox]::Show("Folder A path is invalid."); return }
    if (-not (Test-Path -LiteralPath $FolderB)) { [System.Windows.Forms.MessageBox]::Show("Folder B path is invalid."); return }

    if (-not $OutDir) {
        $OutDir = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        $txtOut.Text = $OutDir
    }
    Ensure-Folder $OutDir

    $FolderA = (Resolve-Path -LiteralPath $FolderA).Path.TrimEnd('\')
    $FolderB = (Resolve-Path -LiteralPath $FolderB).Path.TrimEnd('\')

    Write-Log "Folder A: $FolderA"
    Write-Log "Folder B: $FolderB"
    Write-Log "Output  : $OutDir"
    Write-Log "Recurse : $DoRecurse"
    Write-Log "RelPath : $DoRelPath"
    Write-Log "NameOnly: $DoNameOnly"
    Write-Log "Hash    : $DoHash ($Alg)"
    Write-Log ""

    # Script hash
    $ScriptPath = $PSCommandPath
    $ScriptHash = if ($ScriptPath -and (Test-Path -LiteralPath $ScriptPath)) {
        try { (Get-FileHash -LiteralPath $ScriptPath -Algorithm SHA256).Hash } catch { "ERROR_COMPUTING_HASH" }
    } else { "SCRIPT_PATH_UNAVAILABLE" }

    # Enumerate
    Write-Log "Enumerating files..."
    $filesA = @(Get-FileList $FolderA $DoRecurse)
    $filesB = @(Get-FileList $FolderB $DoRecurse)
    Write-Log ("Folder A files: {0}" -f $filesA.Count)
    Write-Log ("Folder B files: {0}" -f $filesB.Count)
    $progress.Value = 10

    # --- Relative path ---
    $bothPath  = @(); $onlyAPath = @(); $onlyBPath = @()
    if ($DoRelPath) {
        Write-Log "Comparing by relative path..."
        $relA = $filesA | ForEach-Object { Get-RelPath $FolderA $_.FullName }
        $relB = $filesB | ForEach-Object { Get-RelPath $FolderB $_.FullName }

        $setA = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$relA)
        $setB = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$relB)

        $bothPath  = @($setA | Where-Object { $setB.Contains($_) })
        $onlyAPath = @($setA | Where-Object { -not $setB.Contains($_) })
        $onlyBPath = @($setB | Where-Object { -not $setA.Contains($_) })

        $onlyAPath | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_A_ByPath.txt") -Encoding UTF8
        $onlyBPath | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_B_ByPath.txt") -Encoding UTF8
        ($onlyAPath + $onlyBPath) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByPath.txt") -Encoding UTF8

        Write-Log ("ByPath: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $bothPath.Count, $onlyAPath.Count, $onlyBPath.Count)
    } else {
        Write-Log "Skipping relative-path comparison (unchecked)."
    }
    $progress.Value = 30

    # --- Name only ---
    $bothName  = @(); $onlyAName = @(); $onlyBName = @()
    if ($DoNameOnly) {
        Write-Log "Comparing by filename-only..."
        $namesA = $filesA | ForEach-Object Name
        $namesB = $filesB | ForEach-Object Name

        $setA = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$namesA)
        $setB = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$namesB)

        $bothName  = @($setA | Where-Object { $setB.Contains($_) })
        $onlyAName = @($setA | Where-Object { -not $setB.Contains($_) })
        $onlyBName = @($setB | Where-Object { -not $setA.Contains($_) })

        $onlyAName | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_A_ByName.txt") -Encoding UTF8
        $onlyBName | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_B_ByName.txt") -Encoding UTF8
        ($onlyAName + $onlyBName) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByName.txt") -Encoding UTF8

        Write-Log ("ByName: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $bothName.Count, $onlyAName.Count, $onlyBName.Count)
    } else {
        Write-Log "Skipping filename-only comparison (unchecked)."
    }
    $progress.Value = 45

    # --- Hash ---
    $hashesA = New-Object System.Collections.Generic.List[object]
    $hashesB = New-Object System.Collections.Generic.List[object]
    $errA    = New-Object System.Collections.Generic.List[object]
    $errB    = New-Object System.Collections.Generic.List[object]
    $onlyAByHash = @(); $onlyBByHash = @(); $sameDiff = @()

    if ($DoHash) {
        Write-Log "Hashing and comparing content (this may take time)..."
        $total = $filesA.Count + $filesB.Count
        if ($total -lt 1) { throw "No files found to hash." }

        $done = 0

        foreach ($f in $filesA) {
            $done++
            $h = Safe-GetFileHash $f.FullName $Alg
            if ($h) {
                $hashesA.Add([pscustomobject]@{
                    Hash = $h.Hash
                    Rel  = Get-RelPath $FolderA $f.FullName
                    Full = $f.FullName
                    Size = $f.Length
                    LastWriteTimeUtc = $f.LastWriteTimeUtc
                })
            } else {
                $errA.Add([pscustomobject]@{
                    Full  = $f.FullName
                    Rel   = Get-RelPath $FolderA $f.FullName
                    Error = "Hash failed (locked/in use/access denied)."
                })
            }

            if (($done % 50) -eq 0) { Write-Log ("Hashed {0}/{1}..." -f $done, $total) }
            $pct = 45 + [int](45 * $done / $total); if ($pct -gt 90) { $pct = 90 }
            $progress.Value = $pct
        }

        foreach ($f in $filesB) {
            $done++
            $h = Safe-GetFileHash $f.FullName $Alg
            if ($h) {
                $hashesB.Add([pscustomobject]@{
                    Hash = $h.Hash
                    Rel  = Get-RelPath $FolderB $f.FullName
                    Full = $f.FullName
                    Size = $f.Length
                    LastWriteTimeUtc = $f.LastWriteTimeUtc
                })
            } else {
                $errB.Add([pscustomobject]@{
                    Full  = $f.FullName
                    Rel   = Get-RelPath $FolderB $f.FullName
                    Error = "Hash failed (locked/in use/access denied)."
                })
            }

            if (($done % 50) -eq 0) { Write-Log ("Hashed {0}/{1}..." -f $done, $total) }
            $pct = 45 + [int](45 * $done / $total); if ($pct -gt 90) { $pct = 90 }
            $progress.Value = $pct
        }

        $errA | Export-Csv (Join-Path $OutDir "HashErrors_A.csv") -NoTypeInformation -Encoding UTF8
        $errB | Export-Csv (Join-Path $OutDir "HashErrors_B.csv") -NoTypeInformation -Encoding UTF8

        # Build hash -> relpaths maps
        $mapA = @{}; foreach ($row in $hashesA) { if (-not $mapA.ContainsKey($row.Hash)) { $mapA[$row.Hash] = @() }; $mapA[$row.Hash] += $row.Rel }
        $mapB = @{}; foreach ($row in $hashesB) { if (-not $mapB.ContainsKey($row.Hash)) { $mapB[$row.Hash] = @() }; $mapB[$row.Hash] += $row.Rel }

        $hashSetA = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$mapA.Keys)
        $hashSetB = New-Object 'System.Collections.Generic.HashSet[string]' (, [string[]]$mapB.Keys)

        $uniqueHashesA = @($hashSetA | Where-Object { -not $hashSetB.Contains($_) })
        $uniqueHashesB = @($hashSetB | Where-Object { -not $hashSetA.Contains($_) })

        $onlyAByHash = foreach ($hval in $uniqueHashesA) { foreach ($p in ($mapA[$hval] | Sort-Object -Unique)) { [pscustomobject]@{ Hash = $hval; Rel = $p } } }
        $onlyBByHash = foreach ($hval in $uniqueHashesB) { foreach ($p in ($mapB[$hval] | Sort-Object -Unique)) { [pscustomobject]@{ Hash = $hval; Rel = $p } } }

        $onlyAByHash | Export-Csv (Join-Path $OutDir "OnlyIn_A_ByHash.csv") -NoTypeInformation -Encoding UTF8
        $onlyBByHash | Export-Csv (Join-Path $OutDir "OnlyIn_B_ByHash.csv") -NoTypeInformation -Encoding UTF8

        Write-Log ("Hash: HashedA={0} ErrorsA={1} HashedB={2} ErrorsB={3}" -f $hashesA.Count, $errA.Count, $hashesB.Count, $errB.Count)
        Write-Log ("Hash: UniqueContentA={0} UniqueContentB={1}" -f (($onlyAByHash | Measure-Object).Count), (($onlyBByHash | Measure-Object).Count))

        if ($DoSameHashDiffPath) {
            Write-Log "Building 'same hash, different path' report..."
            $sameDiff = foreach ($hval in $mapA.Keys) {
                if ($mapB.ContainsKey($hval)) {
                    $pathsA = ($mapA[$hval] | Sort-Object -Unique)
                    $pathsB = ($mapB[$hval] | Sort-Object -Unique)
                    if (($pathsA -join "|") -ne ($pathsB -join "|")) {
                        [pscustomobject]@{ Hash = $hval; PathsA = ($pathsA -join "; "); PathsB = ($pathsB -join "; ") }
                    }
                }
            }
            $sameDiff | Export-Csv (Join-Path $OutDir "SameHash_DifferentPath.csv") -NoTypeInformation -Encoding UTF8
            Write-Log ("SameHash_DifferentPath rows: {0}" -f (($sameDiff | Measure-Object).Count))
        }
    } else {
        Write-Log "Skipping hash comparison (unchecked)."
    }

    # Summary object
    $SummaryObj = [pscustomobject]@{
        FolderA = $FolderA
        FolderB = $FolderB
        Recurse = $DoRecurse
        FilesA_Enumerated = $filesA.Count
        FilesB_Enumerated = $filesB.Count

        InBoth_ByPath     = if ($DoRelPath)  { $bothPath.Count }  else { $null }
        OnlyInA_ByPath    = if ($DoRelPath)  { $onlyAPath.Count } else { $null }
        OnlyInB_ByPath    = if ($DoRelPath)  { $onlyBPath.Count } else { $null }
        NotInBoth_ByPath  = if ($DoRelPath)  { $onlyAPath.Count + $onlyBPath.Count } else { $null }

        InBoth_ByName     = if ($DoNameOnly) { $bothName.Count }  else { $null }
        OnlyInA_ByName    = if ($DoNameOnly) { $onlyAName.Count } else { $null }
        OnlyInB_ByName    = if ($DoNameOnly) { $onlyBName.Count } else { $null }
        NotInBoth_ByName  = if ($DoNameOnly) { $onlyAName.Count + $onlyBName.Count } else { $null }

        HashedA           = if ($DoHash) { $hashesA.Count } else { $null }
        HashErrorsA       = if ($DoHash) { $errA.Count }    else { $null }
        HashedB           = if ($DoHash) { $hashesB.Count } else { $null }
        HashErrorsB       = if ($DoHash) { $errB.Count }    else { $null }

        UniqueHashesA     = if ($DoHash) { ($onlyAByHash | Measure-Object).Count } else { $null }
        UniqueHashesB     = if ($DoHash) { ($onlyBByHash | Measure-Object).Count } else { $null }

        SameHashDiffPath  = if ($DoHash -and $DoSameHashDiffPath) { ($sameDiff | Measure-Object).Count } else { $null }

        OutputDirectory   = $OutDir
    }

    Write-Log ""
    Write-Log "----- SUMMARY -----"
    ($SummaryObj | Format-List | Out-String).TrimEnd() -split "`r?`n" | ForEach-Object { Write-Log $_ }
    Write-Log "-------------------"

    # Write Summary.txt
    $Header = @()
    $Header += "========================="
    $Header += "Folder Comparison Summary"
    $Header += "========================="
    $Header += "Case Number : $CaseNumber"
    $Header += "Operator    : $Operator"
    $Header += "Run Time    : $($RunTimeLocal.ToString('yyyy-MM-dd HH:mm:ss')) (Local)"
    $Header += "Run Time    : $($RunTimeUtc.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)"
    $Header += "Folder A    : $FolderA"
    $Header += "Folder B    : $FolderB"
    $Header += "Output Dir  : $OutDir"
    $Header += ""

    $Footer = @()
    $Footer += ""
    $Footer += "-----------------------------------------------"
    $Footer += "Tool: $AppName"
    $Footer += "Version: $AppVersion"
    $Footer += "Build Date: $BuildDate"
    $Footer += "Script Hash (SHA256): $ScriptHash"
    $Footer += "-----------------------------------------------"

    ($Header + (($SummaryObj | Format-List | Out-String).TrimEnd()) + $Footer) | Out-File (Join-Path $OutDir "Summary.txt") -Encoding UTF8

    $progress.Value = 100
    Write-Log "Done. Outputs written to: $OutDir"
    [System.Windows.Forms.MessageBox]::Show("Comparison complete.`r`nOutput: $OutDir")
})

# Show form
[void]$form.ShowDialog()
 
