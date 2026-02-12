## Powershell v7
## Install Powershell v7 via PS$ winget install --id Microsoft.Powershell --source winget

 Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Helper functions (UI thread)
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

function Write-Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logBox.AppendText("[$ts] $msg`r`n")
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

# ----------------------------
# GUI Layout
# ----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Compare (Name + Hash) - PS7 Async"
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

# Output
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

# Case / Operator
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

# Options
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

# Run + Progress
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Comparison"
$btnRun.Location = New-Object System.Drawing.Point(12, 305)
$btnRun.Size = New-Object System.Drawing.Size(160, 34)
$btnRun.Font = $font
$form.Controls.Add($btnRun)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(185, 305)
$progress.Size = New-Object System.Drawing.Size(685, 24)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$progress.Anchor = "Top,Left,Right"
$form.Controls.Add($progress)

# Log box
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 350)
$logBox.Size = New-Object System.Drawing.Size(858, 300)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($logBox)

# Default output
$txtOut.Text = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

$AppName    = "RehaCompareGUIv2"
$AppVersion = "3.1"
$BuildDate  = "2026-02-11"

# ----------------------------
# Async engine: RunspacePool + polling timer
# ----------------------------
$sync = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

# Create runspace pool for background work
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1, $iss, $Host)
$pool.ApartmentState = 'MTA'
$pool.Open()

$script = {
    param(
        [hashtable]$p,
        $queue
    )

    function q([int]$pct, [string]$msg) {
        $queue.Enqueue([pscustomobject]@{ Type='progress'; Pct=$pct; Msg=$msg }) | Out-Null
    }

    function Ensure-Folder([string]$Path) {
        if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
    }

function Get-RelPath([string]$Root, [string]$FullName) {
    $root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\')
    $full = $FullName

    # Fast path: same root prefix
    if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($root.Length).TrimStart('\')
    }

    # Different drive/root? fall back to full path
    if ($full.Length -lt 2 -or $root.Length -lt 2 -or $full.Substring(0,1) -ne $root.Substring(0,1)) {
        return $full
    }

    # URI fallback (handles normalization)
    try {
        $rootUri = [System.Uri]::new(($root.TrimEnd('\') + '\'))
        $fullUri = [System.Uri]::new($full)
        $relUri  = $rootUri.MakeRelativeUri($fullUri)
        return [System.Uri]::UnescapeDataString($relUri.ToString()).Replace('/', '\')
    } catch {
        return $full
    }
}


    function Safe-GetFileHash([string]$Path, [string]$Algorithm) {
        try { Get-FileHash -LiteralPath $Path -Algorithm $Algorithm } catch { $null }
    }

    try {
        $FolderA = (Resolve-Path -LiteralPath $p.FolderA).Path.TrimEnd('\')
        $FolderB = (Resolve-Path -LiteralPath $p.FolderB).Path.TrimEnd('\')
        $OutDir  = $p.OutDir
        Ensure-Folder $OutDir

        q 0 "Folder A: $FolderA"
        q 0 "Folder B: $FolderB"
        q 0 "Output  : $OutDir"
        q 0 "Recurse : $($p.DoRecurse)"
        q 0 "RelPath : $($p.DoRelPath)"
        q 0 "NameOnly: $($p.DoNameOnly)"
        q 0 "Hash    : $($p.DoHash) ($($p.Alg))"
        q 0 ""

        q 5 "Enumerating files..."
        $gci = @{ Force=$true; File=$true }
        if ($p.DoRecurse) {
            $filesA = Get-ChildItem -LiteralPath $FolderA -Recurse @gci
            $filesB = Get-ChildItem -LiteralPath $FolderB -Recurse @gci
        } else {
            $filesA = Get-ChildItem -LiteralPath $FolderA @gci
            $filesB = Get-ChildItem -LiteralPath $FolderB @gci
        }
        q 10 ("Folder A files: {0}" -f $filesA.Count)
        q 10 ("Folder B files: {0}" -f $filesB.Count)

        # ByPath
        $bothPath=@(); $onlyAPath=@(); $onlyBPath=@()
        if ($p.DoRelPath) {
            q 15 "Comparing by relative path..."
            [string[]]$relA = $filesA | ForEach-Object { Get-RelPath $FolderA $_.FullName }
            [string[]]$relB = $filesB | ForEach-Object { Get-RelPath $FolderB $_.FullName }

            $setA = [System.Collections.Generic.HashSet[string]]::new($relA)
            $setB = [System.Collections.Generic.HashSet[string]]::new($relB)

            $bothPath  = @($setA | Where-Object { $setB.Contains($_) })
            $onlyAPath = @($setA | Where-Object { -not $setB.Contains($_) })
            $onlyBPath = @($setB | Where-Object { -not $setA.Contains($_) })

            $onlyAPath | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_A_ByPath.txt") -Encoding utf8
            $onlyBPath | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_B_ByPath.txt") -Encoding utf8
            ($onlyAPath + $onlyBPath) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByPath.txt") -Encoding utf8

            q 30 ("ByPath: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $bothPath.Count,$onlyAPath.Count,$onlyBPath.Count)
        } else {
            q 30 "Skipping relative-path comparison (unchecked)."
        }

        # ByName
        $bothName=@(); $onlyAName=@(); $onlyBName=@()
        if ($p.DoNameOnly) {
            q 35 "Comparing by filename-only..."
            [string[]]$namesA = $filesA | ForEach-Object Name
            [string[]]$namesB = $filesB | ForEach-Object Name

            $setA = [System.Collections.Generic.HashSet[string]]::new($namesA)
            $setB = [System.Collections.Generic.HashSet[string]]::new($namesB)

            $bothName  = @($setA | Where-Object { $setB.Contains($_) })
            $onlyAName = @($setA | Where-Object { -not $setB.Contains($_) })
            $onlyBName = @($setB | Where-Object { -not $setA.Contains($_) })

            $onlyAName | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_A_ByName.txt") -Encoding utf8
            $onlyBName | Sort-Object | Out-File (Join-Path $OutDir "OnlyIn_B_ByName.txt") -Encoding utf8
            ($onlyAName + $onlyBName) | Sort-Object -Unique | Out-File (Join-Path $OutDir "NotInBoth_ByName.txt") -Encoding utf8

            q 45 ("ByName: InBoth={0} OnlyInA={1} OnlyInB={2}" -f $bothName.Count,$onlyAName.Count,$onlyBName.Count)
        } else {
            q 45 "Skipping filename-only comparison (unchecked)."
        }

        # Hash
        $hashesA = [System.Collections.Generic.List[object]]::new()
        $hashesB = [System.Collections.Generic.List[object]]::new()
        $errA    = [System.Collections.Generic.List[object]]::new()
        $errB    = [System.Collections.Generic.List[object]]::new()
        $onlyAByHash=@(); $onlyBByHash=@(); $sameDiff=@()

        if ($p.DoHash) {
            q 46 "Hashing and comparing content (this may take time)..."
            $total = $filesA.Count + $filesB.Count
            if ($total -lt 1) { throw "No files found to hash." }

            $done = 0

            foreach ($f in $filesA) {
                $done++
                $h = Safe-GetFileHash $f.FullName $p.Alg
                if ($h) {
                    $hashesA.Add([pscustomobject]@{
                        Hash=$h.Hash; Rel=(Get-RelPath $FolderA $f.FullName); Full=$f.FullName; Size=$f.Length; LastWriteTimeUtc=$f.LastWriteTimeUtc
                    })
                } else {
                    $errA.Add([pscustomobject]@{ Full=$f.FullName; Rel=(Get-RelPath $FolderA $f.FullName); Error="Hash failed (locked/in use/access denied)." })
                }
                if (($done % 50) -eq 0) { q ([Math]::Min(90, 45 + [int](45*$done/$total))) ("Hashed {0}/{1}..." -f $done,$total) }
            }

            foreach ($f in $filesB) {
                $done++
                $h = Safe-GetFileHash $f.FullName $p.Alg
                if ($h) {
                    $hashesB.Add([pscustomobject]@{
                        Hash=$h.Hash; Rel=(Get-RelPath $FolderB $f.FullName); Full=$f.FullName; Size=$f.Length; LastWriteTimeUtc=$f.LastWriteTimeUtc
                    })
                } else {
                    $errB.Add([pscustomobject]@{ Full=$f.FullName; Rel=(Get-RelPath $FolderB $f.FullName); Error="Hash failed (locked/in use/access denied)." })
                }
                if (($done % 50) -eq 0) { q ([Math]::Min(90, 45 + [int](45*$done/$total))) ("Hashed {0}/{1}..." -f $done,$total) }
            }

            $errA | Export-Csv (Join-Path $OutDir "HashErrors_A.csv") -NoTypeInformation -Encoding utf8
            $errB | Export-Csv (Join-Path $OutDir "HashErrors_B.csv") -NoTypeInformation -Encoding utf8

            # maps
            $mapA = @{}; foreach ($row in $hashesA) { if (-not $mapA.ContainsKey($row.Hash)) { $mapA[$row.Hash]=[System.Collections.Generic.List[string]]::new() }; $mapA[$row.Hash].Add($row.Rel) }
            $mapB = @{}; foreach ($row in $hashesB) { if (-not $mapB.ContainsKey($row.Hash)) { $mapB[$row.Hash]=[System.Collections.Generic.List[string]]::new() }; $mapB[$row.Hash].Add($row.Rel) }

            $setHashA = [System.Collections.Generic.HashSet[string]]::new([string[]]$mapA.Keys)
            $setHashB = [System.Collections.Generic.HashSet[string]]::new([string[]]$mapB.Keys)

            $uniqueA = @($setHashA | Where-Object { -not $setHashB.Contains($_) })
            $uniqueB = @($setHashB | Where-Object { -not $setHashA.Contains($_) })

            $onlyAByHash = foreach ($hval in $uniqueA) { foreach ($p1 in ($mapA[$hval] | Sort-Object -Unique)) { [pscustomobject]@{ Hash=$hval; Rel=$p1 } } }
            $onlyBByHash = foreach ($hval in $uniqueB) { foreach ($p1 in ($mapB[$hval] | Sort-Object -Unique)) { [pscustomobject]@{ Hash=$hval; Rel=$p1 } } }

            $onlyAByHash | Export-Csv (Join-Path $OutDir "OnlyIn_A_ByHash.csv") -NoTypeInformation -Encoding utf8
            $onlyBByHash | Export-Csv (Join-Path $OutDir "OnlyIn_B_ByHash.csv") -NoTypeInformation -Encoding utf8

            q 90 ("Hash: HashedA={0} ErrorsA={1} HashedB={2} ErrorsB={3}" -f $hashesA.Count,$errA.Count,$hashesB.Count,$errB.Count)

            if ($p.DoSameHashDiffPath) {
                q 90 "Building 'same hash, different path' report..."
                $sameDiff = foreach ($hval in $mapA.Keys) {
                    if ($mapB.ContainsKey($hval)) {
                        $pa = ($mapA[$hval] | Sort-Object -Unique)
                        $pb = ($mapB[$hval] | Sort-Object -Unique)
                        if (($pa -join "|") -ne ($pb -join "|")) {
                            [pscustomobject]@{ Hash=$hval; PathsA=($pa -join "; "); PathsB=($pb -join "; ") }
                        }
                    }
                }
                $sameDiff | Export-Csv (Join-Path $OutDir "SameHash_DifferentPath.csv") -NoTypeInformation -Encoding utf8
                q 92 ("SameHash_DifferentPath rows: {0}" -f (($sameDiff | Measure-Object).Count))
            }
        } else {
            q 60 "Skipping hash comparison (unchecked)."
        }

        $SummaryObj = [pscustomobject]@{
            FolderA=$FolderA; FolderB=$FolderB; Recurse=$p.DoRecurse
            FilesA_Enumerated=$filesA.Count; FilesB_Enumerated=$filesB.Count
            InBoth_ByPath=if($p.DoRelPath){$bothPath.Count}else{$null}
            OnlyInA_ByPath=if($p.DoRelPath){$onlyAPath.Count}else{$null}
            OnlyInB_ByPath=if($p.DoRelPath){$onlyBPath.Count}else{$null}
            InBoth_ByName=if($p.DoNameOnly){$bothName.Count}else{$null}
            OnlyInA_ByName=if($p.DoNameOnly){$onlyAName.Count}else{$null}
            OnlyInB_ByName=if($p.DoNameOnly){$onlyBName.Count}else{$null}
            HashedA=if($p.DoHash){$hashesA.Count}else{$null}
            HashErrorsA=if($p.DoHash){$errA.Count}else{$null}
            HashedB=if($p.DoHash){$hashesB.Count}else{$null}
            HashErrorsB=if($p.DoHash){$errB.Count}else{$null}
            UniqueHashesA=if($p.DoHash){($onlyAByHash|Measure-Object).Count}else{$null}
            UniqueHashesB=if($p.DoHash){($onlyBByHash|Measure-Object).Count}else{$null}
            SameHashDiffPath=if($p.DoHash -and $p.DoSameHashDiffPath){($sameDiff|Measure-Object).Count}else{$null}
            OutputDirectory=$OutDir
        }

        $header = @(
            "=========================",
            "Folder Comparison Summary",
            "=========================",
            "Case Number : $($p.CaseNumber)",
            "Operator    : $($p.Operator)",
            "Run Time    : $($p.RunTimeLocal.ToString('yyyy-MM-dd HH:mm:ss')) (Local)",
            "Run Time    : $($p.RunTimeUtc.ToString('yyyy-MM-dd HH:mm:ss')) (UTC)",
            "Folder A    : $FolderA",
            "Folder B    : $FolderB",
            "Output Dir  : $OutDir",
            ""
        )

        $footer = @(
            "",
            "-----------------------------------------------",
            "Tool: $($p.AppName)",
            "Version: $($p.AppVersion)",
            "Build Date: $($p.BuildDate)",
            "Script Hash (SHA256): $($p.ScriptHash)",
            "-----------------------------------------------"
        )

        ($header + (($SummaryObj | Format-List | Out-String).TrimEnd()) + $footer) |
            Out-File (Join-Path $OutDir "Summary.txt") -Encoding utf8

        q 95 ""
        q 95 "----- SUMMARY -----"
        (($SummaryObj | Format-List | Out-String).TrimEnd() -split "`r?`n") | ForEach-Object { q 95 $_ }
        q 95 "-------------------"

        $queue.Enqueue([pscustomobject]@{ Type='done'; OutDir=$OutDir }) | Out-Null
    }
    catch {
        $queue.Enqueue([pscustomobject]@{ Type='error'; Error=$_.ToString() }) | Out-Null
    }
}

# state for current async run
$global:psAsync = $null

# UI timer to drain queue + detect completion
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 100
$timer.Add_Tick({
    if ($null -eq $global:psAsync) { return }

    $item = $null   # <-- FIX: create the variable before using [ref]

    while ($sync.TryDequeue([ref]$item)) {
        switch ($item.Type) {
            'progress' {
                if ($item.Msg) { Write-Log $item.Msg }
                $progress.Value = [Math]::Max(0, [Math]::Min(100, [int]$item.Pct))
            }
            'error' {
                $timer.Stop()
                $btnRun.Enabled = $true
                Write-Log ("ERROR: " + $item.Error)
                $progress.Value = 0
                [System.Windows.Forms.MessageBox]::Show("Comparison failed:`r`n`r`n$($item.Error)")
                try { $global:psAsync.PowerShell.EndInvoke($global:psAsync.Handle) | Out-Null } catch {}
                $global:psAsync.PowerShell.Dispose()
                $global:psAsync = $null
            }
            'done' {
                $timer.Stop()
                $btnRun.Enabled = $true
                $progress.Value = 100
                Write-Log "Done. Outputs written to: $($item.OutDir)"
                [System.Windows.Forms.MessageBox]::Show("Comparison complete.`r`nOutput: $($item.OutDir)")
                try { $global:psAsync.PowerShell.EndInvoke($global:psAsync.Handle) | Out-Null } catch {}
                $global:psAsync.PowerShell.Dispose()
                $global:psAsync = $null
            }
        }
    }
})


# ----------------------------
# Button handlers
# ----------------------------
$btnA.Add_Click({ $sel = Select-FolderDialog $txtA.Text; if ($sel) { $txtA.Text = $sel } })
$btnB.Add_Click({ $sel = Select-FolderDialog $txtB.Text; if ($sel) { $txtB.Text = $sel } })
$btnOut.Add_Click({ $sel = Select-FolderDialog $txtOut.Text; if ($sel) { $txtOut.Text = $sel } })

$btnRun.Add_Click({
    if ($global:psAsync) { return }

    $logBox.Clear()
    $progress.Value = 0

    $FolderA = $txtA.Text.Trim()
    $FolderB = $txtB.Text.Trim()
    $OutDir  = $txtOut.Text.Trim()

    if (-not (Test-Path -LiteralPath $FolderA)) { [System.Windows.Forms.MessageBox]::Show("Folder A path is invalid."); return }
    if (-not (Test-Path -LiteralPath $FolderB)) { [System.Windows.Forms.MessageBox]::Show("Folder B path is invalid."); return }

    if (-not $OutDir) {
        $OutDir = Join-Path (Get-Location) ("CompareOutput_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        $txtOut.Text = $OutDir
    }
    Ensure-Folder $OutDir

    # provenance
    $ScriptPath = $PSCommandPath
    $ScriptHash = if ($ScriptPath -and (Test-Path -LiteralPath $ScriptPath)) {
        try { (Get-FileHash -LiteralPath $ScriptPath -Algorithm SHA256).Hash } catch { "ERROR_COMPUTING_HASH" }
    } else { "SCRIPT_PATH_UNAVAILABLE" }

    $p = @{
        FolderA = $FolderA
        FolderB = $FolderB
        OutDir  = $OutDir

        CaseNumber   = $txtCase.Text.Trim()
        Operator     = $txtOp.Text.Trim()
        RunTimeLocal = (Get-Date)
        RunTimeUtc   = (Get-Date).ToUniversalTime()

        DoRecurse          = $chkRecurse.Checked
        DoRelPath          = $chkRelPath.Checked
        DoNameOnly         = $chkNameOnly.Checked
        DoHash             = $chkHash.Checked
        DoSameHashDiffPath = $chkSameHashDiffPath.Checked
        Alg                = $cmbAlg.SelectedItem.ToString()

        AppName    = $AppName
        AppVersion = $AppVersion
        BuildDate  = $BuildDate
        ScriptHash = $ScriptHash
    }

    $btnRun.Enabled = $false
    Write-Log "Starting async comparison..."

    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    $ps.AddScript($script).AddArgument($p).AddArgument($sync) | Out-Null

    $handle = $ps.BeginInvoke()

    $global:psAsync = @{
        PowerShell = $ps
        Handle     = $handle
    }

    $timer.Start()
})

$form.Add_FormClosed({
    # cleanup pool
    try { $timer.Stop() } catch {}
    try { if ($global:psAsync) { $global:psAsync.PowerShell.Stop(); $global:psAsync.PowerShell.Dispose() } } catch {}
    try { $pool.Close(); $pool.Dispose() } catch {}
})

# Show form
[void]$form.ShowDialog()
 
