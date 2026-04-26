# ==============================================================================
# PROJECT: THE ARCHIVIST
# PURPOSE: Universal .msg to Portable HTML/Docx Converter
# AUTHOR:  [jj1eckhardt]
# ==============================================================================

# ==============================================================================
# SECTION 0: INITIALIZATION & IDENTITY
# ==============================================================================
# 1. Manual Identity (Change these for new releases)
$global:ScriptTitle = "THE ARCHIVIST"  # What people see in the UI
$global:ScriptVersion = "v2.0.2"             # The version number
$global:BuildDate = "2026.04.26"         # The build timestamp
#$global:RepoName       = "ArkOS-Utility"      # For GitHub link consistency

# 2. Automated File Identity
# This gets the name of the .ps1 file itself for logging purposes
$global:FName = (Get-Item $PSCommandPath).Basename 

# 3. Initialize Process Globals
$global:MasterPath = ""
$global:ArchivePath = ""
$global:AbortArchive = $false

# ==============================================================================
#  ASCII Art
# ==============================================================================

# Get the filename without the .ps1 extension
$FName = (Get-Item $PSCommandPath).BaseName

# Build the Header Art 
$HeaderArt = @"
  __________________________________________________________________________
 |.                                                                        .|
 |  ' .                                                                . '  |
 |      ' .                                                        . '      |
 |          ' .                  $ScriptTitle                 . '          |
 |              ' .                 $ScriptVersion                 . '              |   
 |                  ' .          $BuildDate            . '                  |
 |                      ' .                        . '                      |
 |                          ' .                . '                          |
 |                              ' .        . '                              |
 |                                  ' .. '                                  |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                                          |
 |                                                              jj1eckhardt |
 |__________________________________________________________________________|

"@

# Display it in the background console immediately
Write-Host $HeaderArt -ForegroundColor Cyan

# ==============================================================================
# SECTION 0: INITIALIZATION
# ==============================================================================
# These three lines MUST run first to load the GUI engine
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()


# Initialize your globals
$global:MasterPath = ""
$global:ArchivePath = ""
$global:IndexContent = ""
$global:AbortArchive = $false


# ==============================================================================
# SECTION 1: UI CONSTRUCTION
# ==============================================================================
$global:Form = New-Object Windows.Forms.Form
$global:Form.Text = "$global:ScriptTitle | $global:ScriptVersion | Build: $global:BuildDate"
$global:Form.Size = New-Object Drawing.Size(650, 540) # Bumped height for the footer
$global:Form.StartPosition = "CenterScreen"
$global:Form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$global:Form.Font = New-Object Drawing.Font("Consolas", 9)
$global:Form.ForeColor = [System.Drawing.Color]::White
$global:Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$global:Form.MaximizeBox = $false
$global:Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide

# --- 1.1. PATH SELECTION GROUP ---
$global:gbPaths = New-Object Windows.Forms.GroupBox
$global:gbPaths.Text = " 1. Setup Source & Target Folder "
$global:gbPaths.Location = New-Object Drawing.Point(20, 20)
$global:gbPaths.Size = New-Object Drawing.Size(595, 120) # Room for all 4 buttons
$global:gbPaths.ForeColor = [System.Drawing.Color]::Cyan

$global:btnMaster = New-Object Windows.Forms.Button
$global:btnMaster.Text = "Sel MASTER (.msg source)"
$global:btnMaster.Location = New-Object Drawing.Point(15, 30)
$global:btnMaster.Size = New-Object Drawing.Size(270, 30)
$global:btnMaster.FlatStyle = "Flat"

$global:btnTarget = New-Object Windows.Forms.Button
$global:btnTarget.Text = "Sel ARCHIVE (destination)"
$global:btnTarget.Location = New-Object Drawing.Point(15, 70)
$global:btnTarget.Size = New-Object Drawing.Size(270, 30)
$global:btnTarget.FlatStyle = "Flat"

$global:btnAbort = New-Object Windows.Forms.Button
$global:btnAbort.Text = "ABORT"
$global:btnAbort.Location = New-Object Drawing.Point(305, 30)
$global:btnAbort.Size = New-Object Drawing.Size(80, 70)
$global:btnAbort.FlatStyle = "Flat"
$global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed
$global:btnAbort.ForeColor = [System.Drawing.Color]::White
$global:btnAbort.Font = New-Object Drawing.Font("Consolas", 10, [Drawing.FontStyle]::Bold)
$global:btnAbort.Enabled = $false

$global:btnReset = New-Object Windows.Forms.Button
$global:btnReset.Text = "RESET SELECTIONS"
$global:btnReset.Location = New-Object Drawing.Point(405, 30)
$global:btnReset.Size = New-Object Drawing.Size(175, 70)
$global:btnReset.FlatStyle = "Flat"
$global:btnReset.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 65)
$global:btnReset.Font = New-Object Drawing.Font("Consolas", 9, [Drawing.FontStyle]::Bold)
$global:btnReset.ForeColor = [System.Drawing.Color]::Cyan

$global:gbPaths.Controls.AddRange(@($global:btnMaster, $global:btnTarget, $global:btnAbort, $global:btnReset))

# --- 1.2. OUTPUT FORMAT GROUP ---
$gbFormat = New-Object Windows.Forms.GroupBox
$gbFormat.Text = " 2. Select Output Format "
$gbFormat.Location = New-Object Drawing.Point(20, 150) # Shifted down to avoid overlap
$gbFormat.Size = New-Object Drawing.Size(285, 120)
$gbFormat.ForeColor = [System.Drawing.Color]::Yellow

$global:rbHTML = New-Object Windows.Forms.RadioButton
$global:rbHTML.Text = "Portable HTML (Universal)"
$global:rbHTML.Location = New-Object Drawing.Point(15, 25)
$global:rbHTML.Checked = $true
$global:rbHTML.AutoSize = $true

$global:rbDOCX = New-Object Windows.Forms.RadioButton
$global:rbDOCX.Text = "MS Word .docx (Embedded)"
$global:rbDOCX.Location = New-Object Drawing.Point(15, 55)
$global:rbDOCX.AutoSize = $true

$global:rbTXT = New-Object Windows.Forms.RadioButton
$global:rbTXT.Text = "Plain Text .txt (Lite)"
$global:rbTXT.Location = New-Object Drawing.Point(15, 85)
$global:rbTXT.AutoSize = $true

$gbFormat.Controls.AddRange(@($global:rbHTML, $global:rbDOCX, $global:rbTXT))

# (Radio buttons at Y: 25, 55, 85)

# --- 1.3. SORTING & ATTACHMENT HANDLING ---
$gbAttach = New-Object Windows.Forms.GroupBox
$gbAttach.Text = " 3. Index / Attachment Handling "
$gbAttach.Location = New-Object Drawing.Point(325, 150)
$gbAttach.Size = New-Object Drawing.Size(290, 120) # Height Matched!
$gbAttach.ForeColor = [System.Drawing.Color]::Lime

$global:chkExtract = New-Object Windows.Forms.CheckBox
$global:chkExtract.Text = "Extract attachments to subfolders"
$global:chkExtract.Location = New-Object Drawing.Point(15, 25) # Top row
$global:chkExtract.AutoSize = $true
$global:chkExtract.Checked = $true

$global:chkIndex = New-Object Windows.Forms.CheckBox
$global:chkIndex.Text = "Generate Master Index Page"
$global:chkIndex.Location = New-Object Drawing.Point(15, 55) # Middle row
$global:chkIndex.AutoSize = $true
$global:chkIndex.Checked = $true

# NEW: Auto-Open Checkbox
$global:chkAutoOpen = New-Object Windows.Forms.CheckBox
$global:chkAutoOpen.Text = "Auto-Open Index when Finished"
$global:chkAutoOpen.Location = New-Object Drawing.Point(15, 85)
$global:chkAutoOpen.Checked = $true # Make it the default
$global:chkAutoOpen.AutoSize = $true

$gbAttach.Controls.AddRange(@($global:chkExtract, $global:chkIndex, $global:chkAutoOpen))

# --- 1.4. ACTION BUTTONS ---

# [ START ] Button (The Workhorse - 80% width)
$global:btnRun = New-Object Windows.Forms.Button
$global:btnRun.Text = "START ARCHIVAL PROCESS"
$global:btnRun.Location = New-Object Drawing.Point(20, 290)
$global:btnRun.Size = New-Object Drawing.Size(470, 45) # Shrunk to make room
$global:btnRun.FlatStyle = "Flat"
$global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
$global:btnRun.Enabled = $false

# [ REBUILD INDEX ] Button (The Auditor - 20% width)
$global:btnRebuild = New-Object Windows.Forms.Button
$global:btnRebuild.Text = "REBUILD INDEX"
$global:btnRebuild.Location = New-Object Drawing.Point(495, 290) # Snapped to the right
$global:btnRebuild.Size = New-Object Drawing.Size(120, 45)
$global:btnRebuild.FlatStyle = "Flat"
$global:btnRebuild.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 45)
$global:btnRebuild.Font = New-Object Drawing.Font("Consolas", 8) # Smaller font to fit
$global:btnRebuild.ForeColor = [System.Drawing.Color]::Cyan

# --- 1.5. DASHBOARD & LOG ---
$global:ProgressBar = New-Object Windows.Forms.ProgressBar
$global:ProgressBar.Location = New-Object Drawing.Point(20, 350)
$global:ProgressBar.Size = New-Object Drawing.Size(595, 15)

$global:lblAction = New-Object Windows.Forms.Label
$global:lblAction.Text = "Ready to Archive..."
$global:lblAction.Location = New-Object Drawing.Point(20, 370)
$global:lblAction.Size = New-Object Drawing.Size(595, 20)

$global:Log = New-Object Windows.Forms.TextBox
$global:Log.Multiline = $true
$global:Log.Location = New-Object Drawing.Point(20, 400)
$global:Log.Size = New-Object Drawing.Size(595, 70)
$global:Log.BackColor = [System.Drawing.Color]::Black
$global:Log.ForeColor = [System.Drawing.Color]::Lime


# --- 1.6. FOOTER & GITHUB LINK ---
$global:lblGitHub = New-Object System.Windows.Forms.LinkLabel
$global:lblGitHub.Text = "$global:FName $global:ScriptVersion | GitHub.com | Click for Updates"
$global:lblGitHub.Location = New-Object Drawing.Point(20, 480) 
$global:lblGitHub.Size = New-Object Drawing.Size(450, 20) # Constrained width
$global:lblGitHub.LinkColor = [System.Drawing.Color]::Gray
$global:lblGitHub.Font = New-Object Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$global:lblGitHub.Add_LinkClicked({ [System.Diagnostics.Process]::Start("https://github.com/jj1eckhardt-beep/MSG-to-DOCX-Batch-Converter") })

# --- 1.7. SUPPORT WIDGET (Aligned to bottom right) ---
$global:lblSupport = New-Object System.Windows.Forms.LinkLabel
# Use the Unicode ID for a Heart [char]0x2764
#$global:lblSupport.Text = "$([char]0x2764) Support"
#$global:lblSupport.Location = New-Object Drawing.Point(545, 480) # Matches GitHub Y-axis
$global:lblSupport.Text = "Support $([char]0x2764)"
$global:lblSupport.Location = New-Object Drawing.Point(545, 480) # Matches GitHub Y-axis
$global:lblSupport.AutoSize = $true
$global:lblSupport.LinkColor = [System.Drawing.Color]::Red
$global:lblSupport.ActiveLinkColor = [System.Drawing.Color]::White
$global:lblSupport.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$global:lblSupport.Add_LinkClicked({ Start-Process "https://ko-fi.com/kofisupporter19535" })


# ==============================================================================
# SECTION 2: CORE FUNCTIONS
# ==============================================================================

# Function 2.1: The HTML Index Header (CSS & Table Start)
function Get-HTMLHeader {
    param($ItemCount)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    return @"
<html><head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: 'Consolas', monospace; background-color: #2D2D2E; color: #FFFFFF; padding: 20px; }
        .stats-bar { color: #00FFFF; font-size: 12px; border-bottom: 1px solid #555; padding-bottom: 10px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; }
        .stats-bar span { flex: 1; }
        .stats-mid { text-align: center; }
        .stats-right { text-align: right; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th { background-color: #483D8B; color: cyan; text-align: left; padding: 12px; border-bottom: 2px solid #555; cursor: pointer; user-select: none; }
        th:hover { background-color: #5A4FCF; }
        td { padding: 10px; border-bottom: 1px solid #444; }
        tr:hover { background-color: #3E3E42; }
        a { color: #FFFF00; text-decoration: none; }
        
        /* --- UPDATED PAPERCLIP CSS --- */
.paperclip { 
    color: #00FF00; 
    text-decoration: none; 
    font-weight: bold; 
    font-size: 18px; 
}
.paperclip:hover { 
    color: #FFFFFF; 
}

    </style>
    <script>
    function sortTable(n) {
        var table = document.querySelector("table");
        var tbody = table.querySelector("tbody");
        if (!tbody) return;
        var rows = Array.from(tbody.rows);
        var dir = table.getAttribute("data-sort-dir-" + n) === "asc" ? "desc" : "asc";
        rows.sort(function(a, b) {
            var valA = a.cells[n].innerText.trim().toLowerCase();
            var valB = b.cells[n].innerText.trim().toLowerCase();
            return dir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
        });
        table.setAttribute("data-sort-dir-" + n, dir);
        rows.forEach(function(row) { tbody.appendChild(row); });
    }
    </script>
</head>
<body>
"<h1>$global:ScriptTitle | $global:ScriptVersion | Build: $global:BuildDate | Master Index</h1>"
        <div class="stats-bar">
        <span><b>Source:</b> $global:MasterPath</span>
        <span class="stats-mid"><b>Destination:</b> $global:ArchivePath</span>
        <span class="stats-right"><b>Count:</b> $ItemCount | <b>Generated:</b> $timestamp</span>
    </div>
    <p style='color:gray; font-size:11px;'><i>Click headers to sort: Date, Subject, or From.</i></p>
    <table>
    <thead>
        <tr>
            <th onclick="sortTable(0)">Date &#9662;</th>
            <th onclick="sortTable(1)">Subject &#9662;</th>
            <th onclick="sortTable(2)">From &#9662;</th>
            <th>&#128206;</th>
        </tr>
    </thead>
    <tbody>
"@
}


# Function 2.2: Add a row to the master index
function Add-IndexRow {
    param($Date, $Subject, $From, $MsgLink, $PdfLink)
    
    # NEW: Cyan Paperclip that links directly to the attachment folder
    $clip = if ($PdfLink) { 
        "<a class='paperclip' href='$PdfLink' title='Open Attachment Folder' style='color:cyan; text-decoration:none;'>&#128206;</a>" 
    }
    else { "" }
    
    return @"
    <tr>
        <td>$Date</td>
        <td><a href='$MsgLink'>$Subject</a></td>
        <td>$From</td>
        <td style='text-align:center;'>$clip</td>
    </tr>
"@
}


# Function 2.3: Update the UI Log Box (The ArkOS Bridge)
function Update-Log {
    param([string]$Message, [bool]$Stamp = $false)
    
    $time = if ($Stamp) { (Get-Date -Format "HH:mm:ss") + " | " } else { "" }
    $line = "`r`n$time$Message"
    
    # Send text to the UI Log box
    $global:Log.AppendText($line)
    
    # Auto-scroll to the bottom
    $global:Log.SelectionStart = $global:Log.Text.Length
    $global:Log.ScrollToCaret()
    
    # Keep the console in sync for debugging
    Write-Host $Message -ForegroundColor Gray

}

# Function 2.4: Universal Progress Bar & Dashboard Updater
function Update-ArchivistProgress {
    param($Count, $Total, $Stopwatch, $FileName)

    # 1. Percent & Time
    $percent = [int](($Count / $Total) * 100)
    $elapsedStr = $Stopwatch.Elapsed.ToString("hh\:mm\:ss")
    $global:ProgressBar.Value = $percent

    # 2. Calculate Items Per Second & ETA 
    $secElapsed = $Stopwatch.Elapsed.TotalSeconds
    $itemsPerSec = if ($secElapsed -gt 0) { $Count / $secElapsed } else { 0 }
    
    $remainStr = "--:--"
    if ($itemsPerSec -gt 0) {
        $secondsLeft = ($Total - $Count) / $itemsPerSec
        $remainStr = [TimeSpan]::FromSeconds($secondsLeft).ToString("mm\:ss")
    }

    # 3. Clean filename for display (Idiot-proof length)
    $displayName = if ($FileName.Length -gt 25) { $FileName.Substring(0, 22) + "..." } else { $FileName }

    # --- 2. The Fixed-Width Template (ArkOS Style) ---
    # {0,3} = Percent | {1,4} = Speed | {2} = Elapsed | {3} = Remaining | {4} = Name
    $template = "Done: {0,3}% | {1,4:N1} Item/s | Time: {2} / {3} | {4}"

    $global:lblAction.Text = $template -f $percent, 
    $([math]::Round($itemsPerSec, 1)), 
    $elapsedStr, # <--- Must match variable above
    $remainStr, 
    $displayName
    
    [System.Windows.Forms.Application]::DoEvents()
}

# Function 2.5: The Auditor (The "Construction Worker")
function Update-MasterIndex {
    Update-Log "GENERATING MASTER INDEX (Safe Audit Mode)..." $true
    
    $indexFile = Join-Path $global:ArchivePath "INDEX.html"
    $content = Get-HTMLHeader # This MUST contain the <meta charset='UTF-8'> and <tbody> tags we discussed
    
    # 1. THE PEACE TREATY: Try to borrow the existing Outlook session first
    $outlook = $null
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    }
    catch {
        try { $outlook = New-Object -ComObject Outlook.Application } catch { $outlook = $null }
    }

    if ($null -eq $outlook) {
        Update-Log "AUDIT ERROR: Outlook engine is blocked or busy." $true
        return
    }

    # 2. Get Items (Simplified)
    $items = Get-ChildItem $global:ArchivePath -Exclude "INDEX.html", "*.dll", "*.bat", "*.ps1"
    $content = Get-HTMLHeader -ItemCount $items.Count

    foreach ($item in $items) {
        $name = $item.BaseName
        $link = if ($item.PSIsContainer) { "$($item.Name)/Message.html" } else { $item.Name }
        
        $displayDate = $item.CreationTime.ToString("yyyy-MM-dd")
        $displaySender = "Archived Item"
        $displaySubject = $item.BaseName
        $pdfLink = "" 

        # 3. THE OUTLOOK PEEK: Find the original .msg
        $cleanSearch = ($name -replace ' -$', '').Trim()
        $sourceFile = Get-ChildItem -Path $global:MasterPath -Filter "*$cleanSearch*.msg" -Recurse | Select-Object -First 1
        
        if ($sourceFile) {
            try {
                $msgObj = $outlook.CreateItemFromTemplate($sourceFile.FullName)
                if ($msgObj.ReceivedTime) { $displayDate = $msgObj.ReceivedTime.ToString("yyyy-MM-dd") }
                if ($msgObj.SenderName) { $displaySender = $msgObj.SenderName }
                if ($msgObj.Subject) { $displaySubject = $msgObj.Subject }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msgObj) | Out-Null
            }
            catch { }
        }

        # 3. Universal Attachment Logic (Auditor Mode)
        if ($item.PSIsContainer) {
            # Look for ANY file that isn't the main Message.html
            $hasAt = Get-ChildItem -Path $item.FullName -Exclude "Message.html" | Select-Object -First 1
            if ($hasAt) { 
                # Link to the folder so they see all attachments
                $pdfLink = "$($item.Name)" 
            }
        }

        # Use the "Brick" from Function 2.2
        # 4. CHANGE THE PARAMETER NAME HERE to match
        $content += Add-IndexRow -Date $displayDate -Subject $displaySubject -From $displaySender -MsgLink $link -PdfLink $pdfLink
        
        [System.Windows.Forms.Application]::DoEvents()
    }

    # 5. SEAL THE FILE (Adding the <tbody> tag and the Total Count)
    $content += "</tbody></table>"
    $content += "<div style='margin-top: 20px; padding: 10px; border-top: 1px solid #555; color: #00FFFF; font-family: Consolas;'>"
    $content += "TOTAL ARCHIVE COUNT: $($items.Count) Items"
    $content += "</div>"
    $content += "<p style='color:gray; font-size:10px;'>Audit Completed: $(Get-Date)</p></body></html>"

    # Force UTF8 to kill the ASCII ghosts
    $content | Out-File $indexFile -Encoding utf8

    
    # 6. LEAVE THE ENGINE ALIVE
    # We don't call $outlook.Quit() here because the main script might still be using it!
    Update-Log "MASTER INDEX SEALED: $($items.Count) items indexed." $true
}


# ==============================================================================
# SECTION 3: FOLDER SELECTION LOGIC
# ==============================================================================

# Function 3.1: Pick the Source (.msg files)
function Set-Path {
    param([string]$Type)
    
    $fb = New-Object Windows.Forms.FolderBrowserDialog
    
    # --- THE ADDITION: Enable the "Make New Folder" button ---
    $fb.ShowNewFolderButton = $true 
    
    if ($fb.ShowDialog($global:Form) -eq [Windows.Forms.DialogResult]::OK) {
        if ($Type -eq "Master") {
            $global:MasterPath = $fb.SelectedPath
            $global:btnMaster.Text = "MASTER: " + (Split-Path $global:MasterPath -Leaf)
            $global:btnMaster.BackColor = [System.Drawing.Color]::DarkGreen
            Update-Log "MASTER SOURCE SET: $global:MasterPath"
        } 
        else {
            $global:ArchivePath = $fb.SelectedPath
            $global:btnTarget.Text = "ARCHIVE: " + (Split-Path $global:ArchivePath -Leaf)
            $global:btnTarget.BackColor = [System.Drawing.Color]::DarkGreen
            Update-Log "ARCHIVE TARGET SET: $global:ArchivePath"
        }

        if ($global:MasterPath -ne "" -and $global:ArchivePath -ne "") {
            $global:btnRun.Enabled = $true
            $global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
            Update-Log "System Ready. Select options and press START."
        }
    }
}


# ==============================================================================
# SECTION 4: THE LOGIC ENGINE (Execution)
# ==============================================================================
# Function 4.1: Path Gatekeeper (Safety Check)
function Test-PathsReady {
    if ([string]::IsNullOrWhiteSpace($global:MasterPath) -or [string]::IsNullOrWhiteSpace($global:ArchivePath)) {
        [Windows.Forms.MessageBox]::Show(
            "Please select both a MASTER source and an ARCHIVE destination folder before starting.",
            "Missing Paths",
            [Windows.Forms.MessageBoxButtons]::OK,
            [Windows.Forms.MessageBoxIcon]::Warning
        )
        return $false
    }
    return $true
}


# Function 4.2: Main Archival Loop
function Invoke-ArchivalProcess {
    if (-not (Test-PathsReady)) { return }

    # 1. Initialize State
    $global:AbortArchive = $false
    $global:btnAbort.Enabled = $true
    $global:btnAbort.Text = "ABORT"
    $global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed
    $global:btnRun.Enabled = $false
    $global:Log.Clear()
    $global:ProgressBar.Value = 0
    Update-Log "INITIALIZING OUTLOOK ENGINE..." $true

    try {
        # 2. Connect to Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $namespace.Logon("", "", $false, $false)

        # 3. Get Files
        $msgFiles = Get-ChildItem -Path $global:MasterPath -Filter *.msg | Sort-Object Name
        if ($msgFiles.Count -eq 0) {
            Update-Log "ERROR: No .msg files found." $true
            return
        }

        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $count = 0

        # --- THE MAIN LOOP ---
        foreach ($file in $msgFiles) {
            if ($global:AbortArchive) {
                Update-Log "!!! ABORTED: Stopping... !!!" $true
                break
            }

            $count++
            Update-ArchivistProgress -Count $count -Total $msgFiles.Count -Stopwatch $timer -FileName $file.Name
            
            $msg = $outlook.CreateItemFromTemplate($file.FullName)
            $safeName = ($file.BaseName -replace '[\\\/\:\*\?\"\<\>\|]', '_').Trim(" -.")
            
            # --- A. HTML BRANCH ---
            if ($global:rbHTML.Checked) {
                $itemFolder = Join-Path $global:ArchivePath $safeName
                $htmlPath = Join-Path $itemFolder "Message.html"
                
                if (Test-Path $htmlPath) { Update-Log "SKIPPING: $safeName"; continue }
                if (-not (Test-Path $itemFolder)) { New-Item -ItemType Directory -Path $itemFolder -Force | Out-Null }
                
                $attachLinks = "<div style='background:#333; padding:10px; border:1px solid #555; margin-bottom:15px; font-family:Consolas;'>"
                $attachLinks += "<b style='color:cyan;'>ATTACHMENTS:</b><br>"
                foreach ($at in $msg.Attachments) {
                    if ($at.FileName -notlike "image*") {
                        $safeAtName = $at.FileName -replace '[\\\/\:\*\?\"\<\>\|]', '_'
                        $at.SaveAsFile((Join-Path $itemFolder $safeAtName))
                        # SIMPLE INLINE STYLE: Forces green color directly on the link
                        $attachLinks += "<p style='margin:5px 0;'><a href='$safeAtName' style='color:cyan; text-decoration:none;'>&#128206; $safeAtName</a></p>"
                    }
                }
                $attachLinks += "</div>"

                $fullHtml = "<html><body>$attachLinks$($msg.HTMLBody)</body></html>"
                $fullHtml | Out-File $htmlPath -Encoding utf8
                Update-Log "Archived (HTML): $safeName"
            } 

            # --- B. DOCX BRANCH ---
            elseif ($global:rbDOCX.Checked) {
                $docxFileName = "$safeName.docx"
                $docxPath = Join-Path $global:ArchivePath $docxFileName
                if (Test-Path $docxPath) { Update-Log "SKIPPING: $docxFileName"; continue }

                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $doc = $word.Documents.Add()

                # Embed Attachments at the TOP
                if ($msg.Attachments.Count -gt 0) {
                    $word.Selection.Font.Bold = $true
                    $word.Selection.TypeText("--- ATTACHED OBJECTS ---")
                    $word.Selection.TypeParagraph()
                    $word.Selection.Font.Bold = $false
                    foreach ($at in $msg.Attachments) {
                        if ($at.FileName -notlike "image*") {
                            $tempFile = Join-Path $env:TEMP $at.FileName
                            $at.SaveAsFile($tempFile)
                            $word.Selection.InlineShapes.AddOLEObject($null, $tempFile, $false, $true, $null, $null, $at.FileName) | Out-Null
                            $word.Selection.TypeParagraph()
                            Remove-Item $tempFile -Force
                        }
                    }
                    $word.Selection.TypeText("--------------------------------------------------")
                    $word.Selection.TypeParagraph()
                }

                $word.Selection.TypeText($msg.Body)
                $doc.SaveAs([string]$docxPath)
                $doc.Close()
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                Update-Log "Archived (DOCX): $docxFileName"
            } 
            # --- C. TXT BRANCH ---
            elseif ($global:rbTXT.Checked) {
                $txtPath = Join-Path $global:ArchivePath "$safeName.txt"
                if (Test-Path $txtPath) { Update-Log "SKIPPING: $safeName"; continue }
                $msg.Body | Out-File -FilePath $txtPath -Encoding utf8
                Update-Log "Archived (TXT): $safeName"
            }

        } # --- END OF FOREACH LOOP ---

        $timer.Stop()
        $finalStatus = if ($global:AbortArchive) { "HALTED" } else { "SUCCESS" }
        Update-Log "$($finalStatus): $($count) files processed in $($timer.Elapsed.ToString('mm\:ss'))" $true
        $global:lblAction.Text = if ($global:AbortArchive) { "Archival HALTED." } else { "Archival Complete!" }

    }
    catch {
        Update-Log "CRITICAL ERROR: $_" $true
    }
    finally {
        $global:btnRun.Enabled = $true
        $global:btnAbort.Enabled = $false
        $global:btnAbort.Text = "ABORT"
        $global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed
        if ($outlook) {
            $outlook.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
    }
}


# ==============================================================================
# SECTION 5: EVENT WIRING (Connecting the Dots)
# ==============================================================================
# Section 5: Wiring
$global:btnMaster.Add_Click({ Set-Path -Type "Master" })
$global:btnTarget.Add_Click({ Set-Path -Type "Target" })
# Wiring the Rebuild button
$global:btnRebuild.Add_Click({ Update-MasterIndex })

# Update the Start button to call the Indexer at the end
$global:btnRun.Add_Click({ 
        Invoke-ArchivalProcess
    
        if ($global:chkIndex.Checked) { 
            Update-MasterIndex 
        
            # NEW: The Auto-Open Handshake
            if ($global:chkAutoOpen.Checked) {
                $indexFile = Join-Path $global:ArchivePath "INDEX.html"
                if (Test-Path $indexFile) { Start-Process $indexFile }
            }
        } 
    })

$global:btnReset.Add_Click({
        # 1. Clear the Global Paths
        $global:MasterPath = ""
        $global:ArchivePath = ""
    
        # 2. Reset the Selection Buttons (Text & Color)
        $global:btnMaster.Text = "Sel MASTER (.msg source)"
        $global:btnMaster.BackColor = [System.Drawing.Color]::Transparent
    
        $global:btnTarget.Text = "Sel ARCHIVE (destination)"
        $global:btnTarget.BackColor = [System.Drawing.Color]::Transparent
    
        # 3. FORCE DEFAULTS: Flip back to HTML Mode
        $global:rbHTML.Checked = $true
        $global:rbDOCX.Checked = $false
    
        # 4. UNLOCK & RE-CHECK the Extraction Box
        $global:chkExtract.Enabled = $true
        $global:chkExtract.Checked = $true
    
        # 5. Reset Action Button & Dashboard
        $global:btnRun.Enabled = $false
        $global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
        $global:ProgressBar.Value = 0
        $global:lblAction.Text = "Ready to Archive..."
    
        Update-Log "--- SYSTEM RESET TO DEFAULTS (HTML MODE) ---" $true
    })
    
# --- ABORT BUTTON WIRING ---
$global:btnAbort.Add_Click({
        # 1. Flip the logic switch
        $global:AbortArchive = $true
    
        # 2. Provide instant UI feedback
        $global:btnAbort.Enabled = $false
        $global:btnAbort.Text = "HALTING"
        $global:btnAbort.BackColor = [System.Drawing.Color]::Gray
    
        Update-Log "!!! ABORT REQUESTED: Cleaning up and stopping... !!!" $true
    })
   
# --- UI INTERACTION: TXT Mode ---
$global:rbTXT.Add_CheckedChanged({
        if ($global:rbTXT.Checked) {
            $global:chkExtract.Checked = $false
            $global:chkExtract.Enabled = $false # Lock it down
            Update-Log "TXT Mode: Lite text only (Extraction disabled)."
        }
    })

# --- UI INTERACTION: Gray out Checkbox for DOCX Mode ---
# 1. Listen to the DOCX Button
$global:rbDOCX.Add_CheckedChanged({
        if ($global:rbDOCX.Checked) {
            $global:chkExtract.Checked = $false
            $global:chkExtract.Enabled = $false # Lock it down
            Update-Log "DOCX Mode: Extraction disabled (PDFs will be EMBEDDED)."
        }
    })

# 2. Listen to the HTML Button
$global:rbHTML.Add_CheckedChanged({
        if ($global:rbHTML.Checked) {
            $global:chkExtract.Enabled = $true  # Unlock it
            $global:chkExtract.Checked = $true  # Default to ON for HTML
            Update-Log "HTML Mode: Extraction re-enabled for folder packaging."
        }
    })

    
# ==============================================================================
# SECTION 6: THE ASSEMBLY & LAUNCH
# ==============================================================================

# --- 6.1: FINAL ASSEMBLY (The Order Matters!) ---
# We clear existing controls to prevent "ghosting" from previous failed attempts
$global:Form.Controls.Clear()

$global:Form.Controls.AddRange(@(
    $global:gbPaths, 
    $gbFormat, 
    $gbAttach, 
    $global:btnRun, 
    $global:btnRebuild, 
    $global:ProgressBar, 
    $global:lblAction, 
    $global:Log, 
    $global:lblGitHub, 
    $global:lblSupport
))

# Launch the app!
$global:Form.ShowDialog()

# --- THE CLEAN EXIT ---
$global:Form.Add_FormClosing({
        # Stop the Outlook engine if it's running
        if ($null -ne $outlook) { 
            $outlook.Quit() 
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
    
        # Optional: Stop any background timers or logs
        Update-Log "Shutting down engine... Goodbye."
    
        # Final cleanup to free memory
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    })

