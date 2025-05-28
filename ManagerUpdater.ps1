Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Load name alias map
$aliasFile = ".\aliases.csv"
$nameAliases = @{}
if (Test-Path $aliasFile) {
    Import-Csv $aliasFile | ForEach-Object {
        $nameAliases[$_.CSVName] = $_.ADName
    }
}

function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path ".\update-log.txt" -Value "$timestamp`t$Message"
}

# UI Setup
$fd = New-Object System.Windows.Forms.OpenFileDialog
$fd.Filter = "CSV (*.csv)|*.csv"
$fd.Title = "Select CSV File"
if ($fd.ShowDialog() -ne 'OK') { return }
$p = $fd.FileName

$form = New-Object Windows.Forms.Form
$form.Text = "ATPC Staff Updater"
$form.Size = New-Object System.Drawing.Size 1000, 600
$form.StartPosition = "CenterScreen"

$g = New-Object Windows.Forms.DataGridView
$g.Location = New-Object System.Drawing.Point 10, 10
$g.Size = New-Object System.Drawing.Size 960, 450
$g.AutoGenerateColumns = $true
$g.SelectionMode = 'FullRowSelect'
$g.MultiSelect = $false
$g.AllowUserToAddRows = $false
$form.Controls.Add($g)

$l = New-Object Windows.Forms.Label
$l.AutoSize = $true
$l.Location = New-Object System.Drawing.Point 10, 470
$form.Controls.Add($l)

function New-Button($text, $x) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $text
    $b.Size = New-Object System.Drawing.Size 150, 30
    $b.Location = New-Object System.Drawing.Point $x, 500
    $b.Enabled = $false
    $form.Controls.Add($b)
    return $b
}

$btn = New-Button "Update Manager" 10
$btn2 = New-Button "Update Title" 170
$fix = New-Button "Correct Manager" 330
$fixTitle = New-Button "Correct Title" 490
$bulkUpdate = New-Button "Bulk Update Matches" 650
$saveCSV = New-Button "Export to CSV" 810

Import-Module ActiveDirectory -EA 0
$x = New-Object System.Collections.Generic.List[Object]

try {
    $data = Import-Csv $p
    foreach ($i in $data) {
        $rawName = "$($i.'First Name') $($i.'Last Name')"
        $n = if ($nameAliases.ContainsKey($rawName)) { $nameAliases[$rawName] } else { $rawName }
        $escapedN = $n -replace "'", "''"

        $m = ""
        if ($i.Supervisor -match "^(.*?),\s*(.*?)$") {
            $rawMgr = "$($Matches[2]) $($Matches[1])"
            $m = if ($nameAliases.ContainsKey($rawMgr)) { $nameAliases[$rawMgr] } else { $rawMgr }
        }
        $escapedM = $m -replace "'", "''"

        $t = $i.'Position Title'

        $cur = ""; $s = "Not Found"; $ct = ""; $ts = "Not Found"

        $u = Get-ADUser -Filter "Name -eq '$escapedN'" -Properties Manager, Title -EA 0
        if ($u) {
            if ($u.Manager) {
                $mgr = Get-ADUser $u.Manager -Properties Name -EA 0
                if ($mgr) {
                    $cur = $mgr.Name
                    $s = if ($cur -eq $m) { "Match" } else { "Mismatch" }
                } else {
                    $s = "Bad Ref"
                }
            } else {
                $s = "Blank"
            }

            $ct = $u.Title
            if ($ct -eq $t) { $ts = "Match" } elseif (!$ct) { $ts = "Blank" } else { $ts = "Mismatch" }
        }

        $x.Add([PSCustomObject]@{
            Employee            = $n
            "Paylocity Manager" = $m
            "AD Manager"        = $cur
            "Manager Status"    = $s
            "Paylocity Title"   = $t
            "AD Title"          = $ct
            "Title Status"      = $ts
        })
    }

    $g.DataSource = $x
    foreach ($c in $g.Columns) { $c.SortMode = 'Automatic' }

    $g.add_CellFormatting({
        param($s, $e)
        $c = $g.Columns[$e.ColumnIndex].Name
        if ($c -eq 'Manager Status' -or $c -eq 'Title Status') {
            switch ($e.Value) {
                "Match"     { $e.CellStyle.BackColor = 'LightGreen' }
                "Mismatch"  { $e.CellStyle.BackColor = 'LightSalmon' }
                "Blank"     { $e.CellStyle.BackColor = 'LightYellow' }
                default     { $e.CellStyle.BackColor = 'LightGray' }
            }
        }
    })

    $btn.Enabled = $btn2.Enabled = $fix.Enabled = $fixTitle.Enabled = $bulkUpdate.Enabled = $saveCSV.Enabled = $true

} catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failure")
}

function upd {
    if ($g.SelectedRows.Count -eq 0) { $l.Text = "No selection"; return }
    $r = $g.SelectedRows[0]
    $emp = $r.Cells["Employee"].Value
    $mgr = $r.Cells["Paylocity Manager"].Value
    $empEsc = $emp -replace "'", "''"
    $mgrEsc = $mgr -replace "'", "''"

    $eu = Get-ADUser -Filter "Name -eq '$empEsc'" -Prop Manager
    $mu = Get-ADUser -Filter "Name -eq '$mgrEsc'"
    if (!$eu) { $l.Text = "No user: $emp"; return }
    if (!$mu) { $l.Text = "No manager: $mgr"; return }

    Set-ADUser $eu -Manager $mu.DistinguishedName
    $r.Cells["Manager Status"].Value = "Match"
    $g.Refresh()
    $l.Text = "$emp manager updated to $mgr"
    Write-Log "$emp manager updated to $mgr"
}

function updTitle {
    if ($g.SelectedRows.Count -eq 0) { $l.Text = "No selection"; return }
    $r = $g.SelectedRows[0]
    $emp = $r.Cells["Employee"].Value
    $title = $r.Cells["Paylocity Title"].Value
    $empEsc = $emp -replace "'", "''"

    $eu = Get-ADUser -Filter "Name -eq '$empEsc'" -Prop Title
    if (!$eu) { $l.Text = "No user: $emp"; return }

    Set-ADUser $eu -Title $title
    $r.Cells["Title Status"].Value = "Match"
    $g.Refresh()
    $l.Text = "$emp title updated to $title"
    Write-Log "$emp title updated to $title"
}

function fixman {
    if ($g.SelectedRows.Count -eq 0) { $l.Text = "No row"; return }
    $r = $g.SelectedRows[0]
    $v = $r.Cells["Paylocity Manager"].Value
    $new = New-Object Windows.Forms.Form
    $new.Text = "Fix Manager"
    $new.Size = New-Object Drawing.Size 400, 150
    $new.StartPosition = "CenterParent"

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "Enter new manager name:"
    $lbl.Location = New-Object Drawing.Point 10, 10
    $lbl.AutoSize = $true
    $new.Controls.Add($lbl)

    $tb = New-Object Windows.Forms.TextBox
    $tb.Text = $v
    $tb.Location = New-Object Drawing.Point 10, 40
    $tb.Size = New-Object Drawing.Size 360, 20
    $new.Controls.Add($tb)

    $ok = New-Object Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Location = New-Object Drawing.Point 150, 70
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $new.AcceptButton = $ok
    $new.Controls.Add($ok)

    if ($new.ShowDialog() -eq 'OK') {
        $updated = $tb.Text
        $r.Cells["Paylocity Manager"].Value = $updated
        $r.Cells["Manager Status"].Value = if ($r.Cells["AD Manager"].Value -eq $updated) { "Match" } else { "Mismatch" }
        $g.Refresh()
        $l.Text = "Manager changed to $updated"
    }
}

function fixTitle {
    if ($g.SelectedRows.Count -eq 0) { $l.Text = "No row"; return }
    $r = $g.SelectedRows[0]
    $v = $r.Cells["Paylocity Title"].Value
    $new = New-Object Windows.Forms.Form
    $new.Text = "Fix Title"
    $new.Size = New-Object Drawing.Size 400, 150
    $new.StartPosition = "CenterParent"

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "Enter new title:"
    $lbl.Location = New-Object Drawing.Point 10, 10
    $lbl.AutoSize = $true
    $new.Controls.Add($lbl)

    $tb = New-Object Windows.Forms.TextBox
    $tb.Text = $v
    $tb.Location = New-Object Drawing.Point 10, 40
    $tb.Size = New-Object Drawing.Size 360, 20
    $new.Controls.Add($tb)

    $ok = New-Object Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Location = New-Object Drawing.Point 150, 70
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $new.AcceptButton = $ok
    $new.Controls.Add($ok)

    if ($new.ShowDialog() -eq 'OK') {
        $updated = $tb.Text
        $r.Cells["Paylocity Title"].Value = $updated
        $r.Cells["Title Status"].Value = if ($r.Cells["AD Title"].Value -eq $updated) { "Match" } else { "Mismatch" }
        $g.Refresh()
        $l.Text = "Title changed to $updated"
    }
}

function bulkUpdateAll {
    $count = 0
    foreach ($r in $g.Rows) {
        if ($r.Cells["Manager Status"].Value -eq "Mismatch" -or $r.Cells["Title Status"].Value -eq "Mismatch") {
            $count++
        }
    }
    if ($count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("All records are up-to-date.", "No Action")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Update $count mismatched records in AD?", "Confirm", "YesNo", "Warning")
    if ($confirm -ne 'Yes') { return }

    foreach ($r in $g.Rows) {
        $emp = $r.Cells["Employee"].Value
        $empEsc = $emp -replace "'", "''"

        if ($r.Cells["Manager Status"].Value -eq "Mismatch") {
            $mgr = $r.Cells["Paylocity Manager"].Value
            $mgrEsc = $mgr -replace "'", "''"
            $eu = Get-ADUser -Filter "Name -eq '$empEsc'" -Prop Manager
            $mu = Get-ADUser -Filter "Name -eq '$mgrEsc'"
            if ($eu -and $mu) {
                Set-ADUser $eu -Manager $mu.DistinguishedName
                $r.Cells["Manager Status"].Value = "Match"
                Write-Log "$emp manager updated to $mgr (bulk)"
            }
        }

        if ($r.Cells["Title Status"].Value -eq "Mismatch") {
            $title = $r.Cells["Paylocity Title"].Value
            $eu = Get-ADUser -Filter "Name -eq '$empEsc'" -Prop Title
            if ($eu) {
                Set-ADUser $eu -Title $title
                $r.Cells["Title Status"].Value = "Match"
                Write-Log "$emp title updated to $title (bulk)"
            }
        }
    }

    $g.Refresh()
    $l.Text = "Bulk update complete."
}

function exportToCSV {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveDialog.Title = "Save Grid Data"
    $saveDialog.FileName = "AD_Report.csv"
    if ($saveDialog.ShowDialog() -eq 'OK') {
        $export = foreach ($r in $g.Rows) {
            [PSCustomObject]@{
                Employee            = $r.Cells["Employee"].Value
                "Paylocity Manager" = $r.Cells["Paylocity Manager"].Value
                "AD Manager"        = $r.Cells["AD Manager"].Value
                "Manager Status"    = $r.Cells["Manager Status"].Value
                "Paylocity Title"   = $r.Cells["Paylocity Title"].Value
                "AD Title"          = $r.Cells["AD Title"].Value
                "Title Status"      = $r.Cells["Title Status"].Value
            }
        }
        $export | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exported to $($saveDialog.FileName)", "Done")
    }
}

$btn.Add_Click({ upd })
$btn2.Add_Click({ updTitle })
$fix.Add_Click({ fixman })
$fixTitle.Add_Click({ fixTitle })
$bulkUpdate.Add_Click({ bulkUpdateAll })
$saveCSV.Add_Click({ exportToCSV })

[void]$form.ShowDialog()
