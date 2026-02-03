Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Helper functions
# ----------------------------
function Select-FolderDialog([string]$InitialPath = "") {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($InitialPath -and (Test-Path $InitialPath)) { $dlg.SelectedPath = $InitialPath }
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return $null
}

function Ensure-Folder([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
}

function Get-RelPath([string]$Root, [string]$FullName) {
    $root = (Resolve-Path $Root).Path.TrimEnd('\')
    return $FullName.Substring($root.Length).TrimStart('\')
}

function Write-Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logBox.AppendText("[$ts] $msg`r`n")
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}
function Write-SummaryObjectToLog([pscustomobject]$obj) {
    Write-Log "----- SUMMARY -----"
    ($obj | Format-List | Out-String).TrimEnd() -split "`r?`n" | ForEach-Object { Write-Log $_ }
    Write-Log "-------------------"
}

function Get-FileList([string]$Root, [bool]$Recurse) {
    $root = (Resolve-Path $Root).Path.TrimEnd('\')
    if ($Recurse) {
        return Get-ChildItem $root -Recurse -File -Force
    } else {
        return Get-ChildItem $root -File -Force
    }
}

function Safe-GetFileHash([string]$Path, [string]$Algorithm) {
    try {
        return Get-FileHash -Path $Path -Algorithm $Algorithm
    } catch {
        return $null
    }
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
$form.Controls.Add($txtA)
$txtA.Anchor = "Top,Left,Right"

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
$form.Controls.Add($txtB)
$txtB.Anchor = "Top,Left,Right"

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
$form.Controls.Add($txtOut)
$txtOut.Anchor = "Top,Left,Right"

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
$form.Controls.Add($grp)
$grp.Anchor = "Top,Left,Right"

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
$cmbAlg.Items.AddRange(@("SHA256","SHA1","MD5"))
$cmbAlg.SelectedItem = "SHA256"
$cmbAlg.Font = $font
$grp.Controls.Add($cmbAlg)

# Run button + Progress
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Comparison"
$btnRun.Location = New-Object System.Drawing.Point(12, 305)
$btnRun.Size = New-Object System.Drawing.Size(160, 34)
$btnRun.Font = $font
$form.Controls.Add($btnRun)
$btnRun.Anchor = "Top,Left"

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(185, 305)
$progress.Size = New-Object System.Drawing.Size(685, 24)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)
$progress.Anchor = "Top,Left,Right"

# Log window
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 350)
$logBox.Size = New-Object System.Drawing.Size(858, 300)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)
$logBox.Anchor = "Top,Bottom,Left,Right"
$progress.Anchor = "Top,Left,Right"
$btnRun.Anchor = "Top,Left"
$grp.Anchor = "Top,Left,Right"
$txtA.Anchor = "Top,Left,Right"
$txtB.Anchor = "Top,Left,Right"
$txtOut.Anchor = "Top,Left,Right"

# Application metadata
$AppName    = "RehaCompareGUI"
$AppVersion = "2.0"
$BuildDate  = "2026-02-03"

# ----------------------------
# Button handlers
# ----------------------------
$btnA.Add_Click({
    $sel = Select-FolderDialog $txtA.Text
    if ($sel) { $txtA.Text = $sel }
})

$btnB.Add_Click({
    $sel = Select-FolderDialog $txtB.Text
    if ($sel) { $txtB.Text = $sel }
})

$btnOut.Add_Click({
    $sel = Select-FolderDialog $txtOut.Text
    if ($sel) { $txtOut.Text = $sel }
})

# ----------------------------
# Run logic
# ----------------------------
$btnRun.Add_Click({

    $progress.Value = 0
    $logBox.Clear()

$CaseNumber  = $txtCase.Text.Trim()
$Operator    = $txtOp.Text.Trim()
$RunTimeLocal = Get-Date
$RunTimeUtc   = (Get-Date).ToUniversalTime()
    $FolderA = $txtA.Text.Trim()
    $FolderB = $txtB.Text.Trim()
    $OutDir  = $txtOut.Text.Trim()
    $DoRecurse = $chkRecurse.Checked
    $DoRelPath = $chkRelPath.Checked
    $DoNameOnly = $chkNameOnly.Checked
    $DoHash = $chkHash.Checked
    $DoSameHashDiffPath = $chkSameHashDiffPath.Checked
    $Alg = $cmbAlg.SelectedItem.ToString()

    if (-not (Test-Path $FolderA)) { [System.Windows.Forms.MessageBox]::Show("Folder A path is invalid."); return }
    if (-not (Test-Path $FolderB)) { [System.Windows.Forms.MessageBox]::Show("Folder B path is invalid."); return }
    if (-not $OutDir) {
        $OutDir = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        $txtOut.Text = $OutDir
    }

    Ensure-Folder $OutDir

    $FolderA = (Resolve-Path $FolderA).Path.TrimEnd('\')
    $FolderB = (Resolve-Path $FolderB).Path.TrimEnd('\')

    Write-Log "Folder A: $FolderA"
    Write-Log "Folder B: $FolderB"
    Write-Log "Output  : $OutDir"
    Write-Log "Recurse : $DoRecurse"
    Write-Log "RelPath : $DoRelPath"
    Write-Log "NameOnly: $DoNameOnly"
    Write-Log "Hash    : $DoHash ($Alg)"
    Write-Log ""

    # Gather files
    Write-Log "Enumerating files..."
    $filesA = Get-FileList $FolderA $DoRecurse
    $filesB = Get-FileList $FolderB $DoRecurse
    Write-Log ("Folder A files: {0}" -f $filesA.Count)
    Write-Log ("Folder B files: {0}" -f $filesB.Count)
    $progress.Value = 10

# Compute script hash for provenance
$ScriptPath = $PSCommandPath

$ScriptHash = if ($ScriptPath -and (Test-Path $ScriptPath)) {
    try {
        (Get-FileHash -Path $ScriptPath -Algorithm SHA256).Hash
    } catch {
        "ERROR_COMPUTING_HASH"
    }
} else {
    "SCRIPT_PATH_UNAVAILABLE"
}


    # --- Relative path comparison ---
    if ($DoRelPath) {
        Write-Log "Comparing by relative path..."
        $relA = $filesA | ForEach-Object { Get-RelPath $FolderA $_.FullName } | Sort-Object -Unique
        $relB = $filesB | ForEach-Object { Get-RelPath $FolderB $_.FullName } | Sort-Object -Unique

        $diff = Compare-Object -ReferenceObject $relA -DifferenceObject $relB -IncludeEqual
        $onlyA = $diff | Where-Object SideIndicator -eq "<=" | Select-Object -ExpandProperty InputObject
        $onlyB = $diff | Where-Object SideIndicator -eq "=>" | Select-Object -ExpandProperty InputObject
        $both  = $diff | Where-Object SideIndicator -eq "==" | Select-Object -ExpandProperty InputObject

        $onlyA | Out-File (Join-Path $OutDir "OnlyIn_A_ByPath.txt") -Encoding UTF8
        $onlyB | Out-File (Join-Path $OutDir "OnlyIn_B_ByPath.txt") -Encoding UTF8
        ($onlyA + $onlyB) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByPath.txt") -Encoding UTF8

        Write-Log ("ByPath: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $both.Count, $onlyA.Count, $onlyB.Count)
    } else {
        Write-Log "Skipping relative-path comparison (unchecked)."
    }
    $progress.Value = 30

    # --- Filename-only comparison ---
    if ($DoNameOnly) {
        Write-Log "Comparing by filename-only..."
        $namesA = $filesA | Select-Object -ExpandProperty Name | Sort-Object -Unique
        $namesB = $filesB | Select-Object -ExpandProperty Name | Sort-Object -Unique

        $diffN = Compare-Object -ReferenceObject $namesA -DifferenceObject $namesB -IncludeEqual
        $onlyA = $diffN | Where-Object SideIndicator -eq "<=" | Select-Object -ExpandProperty InputObject
        $onlyB = $diffN | Where-Object SideIndicator -eq "=>" | Select-Object -ExpandProperty InputObject
        $both  = $diffN | Where-Object SideIndicator -eq "==" | Select-Object -ExpandProperty InputObject

        $onlyA | Out-File (Join-Path $OutDir "OnlyIn_A_ByName.txt") -Encoding UTF8
        $onlyB | Out-File (Join-Path $OutDir "OnlyIn_B_ByName.txt") -Encoding UTF8
        ($onlyA + $onlyB) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByName.txt") -Encoding UTF8

        Write-Log ("ByName: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $both.Count, $onlyA.Count, $onlyB.Count)
    } else {
        Write-Log "Skipping filename-only comparison (unchecked)."
    }
    $progress.Value = 45

    # --- Hash comparison ---
    if ($DoHash) {
        Write-Log "Hashing and comparing content (this may take time)..."

        $total = $filesA.Count + $filesB.Count
        if ($total -lt 1) { throw "No files found to hash." }

        $hashesA = New-Object System.Collections.Generic.List[object]
        $hashesB = New-Object System.Collections.Generic.List[object]
        $errA = New-Object System.Collections.Generic.List[object]
        $errB = New-Object System.Collections.Generic.List[object]

        $i = 0

        foreach ($f in $filesA) {
            $i++
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
            $pct = 45 + [int](45 * ($i / $total))
            if ($pct -gt 90) { $pct = 90 }
            $progress.Value = $pct
            if (($i % 50) -eq 0) { Write-Log ("Hashed {0}/{1}..." -f $i, $total) }
        }

        foreach ($f in $filesB) {
            $i++
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
            $pct = 45 + [int](45 * ($i / $total))
            if ($pct -gt 90) { $pct = 90 }
            $progress.Value = $pct
            if (($i % 50) -eq 0) { Write-Log ("Hashed {0}/{1}..." -f $i, $total) }
        }

        # Write error logs
        $errA | Export-Csv (Join-Path $OutDir "HashErrors_A.csv") -NoTypeInformation -Encoding UTF8
        $errB | Export-Csv (Join-Path $OutDir "HashErrors_B.csv") -NoTypeInformation -Encoding UTF8

        # Compare hash sets
        $uniqueA = Compare-Object $hashesA.Hash $hashesB.Hash | Where-Object SideIndicator -eq "<=" | Select-Object -ExpandProperty InputObject
        $uniqueB = Compare-Object $hashesA.Hash $hashesB.Hash | Where-Object SideIndicator -eq "=>" | Select-Object -ExpandProperty InputObject

        $onlyAByHash = $hashesA | Where-Object { $uniqueA -contains $_.Hash } | Sort-Object Hash, Rel
        $onlyBByHash = $hashesB | Where-Object { $uniqueB -contains $_.Hash } | Sort-Object Hash, Rel

        $onlyAByHash | Export-Csv (Join-Path $OutDir "OnlyIn_A_ByHash.csv") -NoTypeInformation -Encoding UTF8
        $onlyBByHash | Export-Csv (Join-Path $OutDir "OnlyIn_B_ByHash.csv") -NoTypeInformation -Encoding UTF8

        Write-Log ("Hash: HashedA={0} ErrorsA={1} HashedB={2} ErrorsB={3}" -f $hashesA.Count, $errA.Count, $hashesB.Count, $errB.Count)
        Write-Log ("Hash: UniqueContentA={0} UniqueContentB={1}" -f $onlyAByHash.Count, $onlyBByHash.Count)

        # Same hash, different path report
        if ($DoSameHashDiffPath) {
            Write-Log "Building 'same hash, different path' report..."
            $common = ($hashesA.Hash | Sort-Object -Unique) | Where-Object { $hashesB.Hash -contains $_ }

            $sameDiff = foreach ($h in $common) {
                $pathsA = ($hashesA | Where-Object Hash -eq $h | Select-Object -ExpandProperty Rel) | Sort-Object -Unique
                $pathsB = ($hashesB | Where-Object Hash -eq $h | Select-Object -ExpandProperty Rel) | Sort-Object -Unique
                if (($pathsA -join "|") -ne ($pathsB -join "|")) {
                    [pscustomobject]@{
                        Hash   = $h
                        PathsA = ($pathsA -join "; ")
                        PathsB = ($pathsB -join "; ")
                    }
                }
            }

            $sameDiff | Export-Csv (Join-Path $OutDir "SameHash_DifferentPath.csv") -NoTypeInformation -Encoding UTF8
            Write-Log ("SameHash_DifferentPath rows: {0}" -f (($sameDiff | Measure-Object).Count))
        }
    } else {
        Write-Log "Skipping hash comparison (unchecked)."
    }

    # Summary
    $progress.Value = 95
#    $summaryPath = Join-Path $OutDir "Summary.txt"
#    $summary = @()
#    $summary += "Folder A: $FolderA"
#    $summary += "Folder B: $FolderB"
#    $summary += "Output  : $OutDir"
#    $summary += "Recurse : $DoRecurse"
#    $summary += "RelPath : $DoRelPath"
#    $summary += "NameOnly: $DoNameOnly"
#    $summary += "Hash    : $DoHash ($Alg)"
#    $summary += ""
#    $summary += "Generated files in output directory."
#
#    $summary | Out-File $summaryPath -Encoding UTF8
#   $progress.Value = 100
    Write-Log ""
# ---------- BUILD SUMMARY OBJECT ----------
$SummaryObj = [pscustomobject]@{
    FolderA = $FolderA
    FolderB = $FolderB
    Recurse = $DoRecurse

    FilesA_Enumerated = $filesA.Count
    FilesB_Enumerated = $filesB.Count

    InBoth_ByPath     = if ($DoRelPath) { $bothPath.Count } else { $null }
    OnlyInA_ByPath    = if ($DoRelPath) { $onlyAPath.Count } else { $null }
    OnlyInB_ByPath    = if ($DoRelPath) { $onlyBPath.Count } else { $null }
    NotInBoth_ByPath  = if ($DoRelPath) { $onlyAPath.Count + $onlyBPath.Count } else { $null }

    InBoth_ByName     = if ($DoNameOnly) { $bothName.Count } else { $null }
    OnlyInA_ByName    = if ($DoNameOnly) { $onlyAName.Count } else { $null }
    OnlyInB_ByName    = if ($DoNameOnly) { $onlyBName.Count } else { $null }
    NotInBoth_ByName  = if ($DoNameOnly) { $onlyAName.Count + $onlyBName.Count } else { $null }

    HashedA           = if ($DoHash) { $hashesA.Count } else { $null }
    HashErrorsA       = if ($DoHash) { $errA.Count } else { $null }
    HashedB           = if ($DoHash) { $hashesB.Count } else { $null }
    HashErrorsB       = if ($DoHash) { $errB.Count } else { $null }

    UniqueHashesA     = if ($DoHash) { $onlyAByHash.Count } else { $null }
    UniqueHashesB     = if ($DoHash) { $onlyBByHash.Count } else { $null }

    SameHashDiffPath  = if ($DoSameHashDiffPath) {
                            ($sameDiff | Measure-Object).Count
                        } else { $null }

    OutputDirectory   = $OutDir
}
# ---------- WRITE SUMMARY ----------
Write-Log "----- SUMMARY -----"
($SummaryObj | Format-List | Out-String).TrimEnd() -split "`r?`n" |
    ForEach-Object { Write-Log $_ }
Write-Log "-------------------"

# ---------- WRITE SUMMARY FILE WITH HEADER ----------
$AppName = "RehaCompareGUIv2"

# Make sure these exist earlier in the click handler:
# $CaseNumber, $Operator, $RunTimeLocal, $RunTimeUtc

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

$SummaryBody = ($SummaryObj | Format-List | Out-String).TrimEnd()
$FullSummary = $Header + $SummaryBody

$FullSummary | Out-File (Join-Path $OutDir "Summary.txt") -Encoding UTF8
# ---------- APP FOOTER ----------
$Footer = @()
$Footer += ""
$Footer += "-----------------------------------------------"
$Footer += "Tool: $AppName"
$Footer += "Version: $AppVersion"
$Footer += "Build Date: $BuildDate"
$Footer += "Script Hash (SHA256): $ScriptHash"
$Footer += "-----------------------------------------------"

$Footer | Out-File (Join-Path $OutDir "Summary.txt") -Encoding UTF8 -Append


    Write-Log "Done. Outputs written to: $OutDir"
    [System.Windows.Forms.MessageBox]::Show("Comparison complete.`r`nOutput: $OutDir")
})

# Default output folder suggestion
$txtOut.Text = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

# Show form
[void]$form.ShowDialog()
