# ==============================================================================
# PROJECT: THE ARCHIVIST
# PURPOSE: Universal .msg to Portable HTML/Docx/txt Converter
# AUTHOR:  [jj1eckhardt]
# ==============================================================================


# ==============================================================================
# SECTION 0: INITIALIZATION & IDENTITY
# ==============================================================================
# 1. Manual Identity (Change these for new releases)
$global:ScriptTitle1 = "THE ARCHIVIST"   # What people see in the UI
$global:Cvers = "v2.2.1"                 # The version number
$global:BuildDate = "2026.06.07"         # The build timestamp
$global:RepoName = "The-Archivist"       # For GitHub link consistency
$global:RepoOwner = "jj1eckhardt-beep"

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

# Build the Header Art 
$HeaderArt = @"
  __________________________________________________________________________
 |.                                                                        .|
 |  ' .                                                                . '  |
 |      ' .                                                        . '      |
 |          ' .                  $ScriptTitle1                 . '          |
 |              ' .                 $Cvers                 . '              |   
 |                  ' .           $BuildDate           . '                  |
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
$global:Form.Text = "$global:ScriptTitle1 | $global:Cvers | Build: $global:BuildDate"
$global:Form.Size = New-Object Drawing.Size(650, 540) # Bumped height for the footer
$global:Form.StartPosition = "CenterScreen"
$global:Form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$global:Form.Font = New-Object Drawing.Font("Consolas", 9)
$global:Form.ForeColor = [System.Drawing.Color]::White
$global:Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$global:Form.MaximizeBox = $false
$global:Form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide

# --- 1.1. PATH SELECTION GROUP (WITH NESTED SOURCE BOX) ---
$global:gbPaths = New-Object Windows.Forms.GroupBox
$global:gbPaths.Text = " 1. Setup Source & Target Folder "
$global:gbPaths.Location = New-Object Drawing.Point(20, 20)
$global:gbPaths.Size = New-Object Drawing.Size(595, 120)
$global:gbPaths.ForeColor = [System.Drawing.Color]::Cyan

# [ MASTER ] & [ ARCHIVE ] Buttons
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

# --- NESTED SOURCE TYPE BOX (This creates the internal lines) ---
$global:gbSourceType = New-Object Windows.Forms.GroupBox
$global:gbSourceType.Text = " Source Type"
$global:gbSourceType.Location = New-Object Drawing.Point(306, 24)
$global:gbSourceType.Size = New-Object Drawing.Size(135, 77)
$global:gbSourceType.ForeColor = [System.Drawing.Color]::White
$global:gbSourceType.Font = New-Object Drawing.Font("Consolas", 8)

$global:rbSourceFolder = New-Object Windows.Forms.RadioButton
$global:rbSourceFolder.Text = "Folder"
$global:rbSourceFolder.Location = New-Object Drawing.Point(10, 22)
$global:rbSourceFolder.Checked = $true
$global:rbSourceFolder.AutoSize = $true

$global:rbSourcePST = New-Object Windows.Forms.RadioButton
$global:rbSourcePST.Text = ".pst Archive"
$global:rbSourcePST.Location = New-Object Drawing.Point(10, 47)
$global:rbSourcePST.AutoSize = $true

# Add RadioButtons to the NESTED box
$global:gbSourceType.Controls.AddRange(@($global:rbSourceFolder, $global:rbSourcePST))

# [ RESET ] Button
$global:btnReset = New-Object Windows.Forms.Button
$global:btnReset.Text = "RESET SELECTIONS"
$global:btnReset.Location = New-Object Drawing.Point(450, 30)
$global:btnReset.Size = New-Object Drawing.Size(135, 70)
$global:btnReset.FlatStyle = "Flat"
$global:btnReset.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 65)
$global:btnReset.ForeColor = [System.Drawing.Color]::Cyan
$global:btnReset.Font = New-Object Drawing.Font("Consolas", 8, [Drawing.FontStyle]::Bold)

# Add everything to the MAIN box
$global:gbPaths.Controls.AddRange(@($global:btnMaster, $global:btnTarget, $global:gbSourceType, $global:btnReset))

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
$global:rbDOCX.Text = "MS Word .docx (Embedded BETA)"
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

# --- 1.4. ACTION BUTTONS (REWORKED) ---
# [ START ] Button (65% width)
$global:btnRun = New-Object Windows.Forms.Button
$global:btnRun.Text = "START ARCHIVAL PROCESS"
$global:btnRun.Location = New-Object Drawing.Point(20, 290)
$global:btnRun.Size = New-Object Drawing.Size(285, 45)
$global:btnRun.FlatStyle = "Flat"
$global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
$global:btnRun.Enabled = $false

# [ ABORT ] Button (Relocated - Now a square next to Start)
$global:btnAbort = New-Object Windows.Forms.Button
$global:btnAbort.Text = "ABORT"
$global:btnAbort.Location = New-Object Drawing.Point(325, 290)
$global:btnAbort.Size = New-Object Drawing.Size(135, 45)
$global:btnAbort.FlatStyle = "Flat"
$global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed
$global:btnAbort.ForeColor = [System.Drawing.Color]::White
$global:btnAbort.Font = New-Object Drawing.Font("Consolas", 10, [Drawing.FontStyle]::Bold)
$global:btnAbort.Enabled = $false

# [ REBUILD INDEX ] Button (Remains on the right)
$global:btnRebuild = New-Object Windows.Forms.Button
$global:btnRebuild.Text = "REBUILD INDEX"
$global:btnRebuild.Location = New-Object Drawing.Point(470, 290)
$global:btnRebuild.Size = New-Object Drawing.Size(145, 45)
$global:btnRebuild.FlatStyle = "Flat"
$global:btnRebuild.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 45)
$global:btnRebuild.Font = New-Object Drawing.Font("Consolas", 8)
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
$global:lblGitHub.Text = "$global:FName $global:Cvers | GitHub.com | Click for Updates"
$global:lblGitHub.Location = New-Object Drawing.Point(20, 480) 
$global:lblGitHub.Size = New-Object Drawing.Size(450, 20) # Constrained width
$global:lblGitHub.LinkColor = [System.Drawing.Color]::Gray
$global:lblGitHub.Font = New-Object Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$global:lblGitHub.Add_LinkClicked({ [System.Diagnostics.Process]::Start("https://github.com/$($global:RepoOwner)/$($global:RepoName)/releases/latest") })

# --- 1.7. SUPPORT WIDGET (Symmetrical Static + Link) ---
# 1. The Static Text (Gray)
$global:lblSupportText = New-Object System.Windows.Forms.Label
$global:lblSupportText.Text = "Support"
$global:lblSupportText.Location = New-Object Drawing.Point(550, 480) # Matches GitHub Y-axis
$global:lblSupportText.AutoSize = $true
$global:lblSupportText.ForeColor = [System.Drawing.Color]::Gray
$global:lblSupportText.Font = New-Object Drawing.Font("Segoe UI Symbol", 9, [System.Drawing.FontStyle]::Italic)

# 2. The Heart Link (Red)
$global:lblSupportLink = New-Object System.Windows.Forms.LinkLabel
# Using the Unicode Heart [char]0x2764
$global:lblSupportLink.Text = "$([char]0x2764)"
$global:lblSupportLink.Location = New-Object Drawing.Point(600, 480) # Positioned just after the text
$global:lblSupportLink.AutoSize = $true
$global:lblSupportLink.LinkColor = [System.Drawing.Color]::Red
$global:lblSupportLink.ActiveLinkColor = [System.Drawing.Color]::White
$global:lblSupportLink.Font = New-Object Drawing.Font("Segoe UI Symbol", 10, [System.Drawing.FontStyle]::Bold)
$global:lblSupportLink.LinkBehavior = [System.Windows.Forms.LinkBehavior]::NeverUnderline
$global:lblSupportLink.Add_LinkClicked({ Start-Process "https://ko-fi.com/kofisupporter19535" })
# 1. Create the ToolTip object
$global:toolTipSupport = New-Object System.Windows.Forms.ToolTip
$global:toolTipSupport.InitialDelay = 500  # Delay in milliseconds before showing
$global:toolTipSupport.AutoPopDelay = 5000 # How long it stays visible
$global:toolTipSupport.ToolTipTitle = "Support the Project"

# 2. Attach it to your Heart label
$global:toolTipSupport.SetToolTip($global:lblSupportLink, "Click here to support via Ko-fi!")


# ==============================================================================
# SECTION 2: CORE FUNCTIONS
# ==============================================================================

# Function 2.1: Updated: The HTML Index Header (CSS & Table Start)
function Get-HTMLHeader { 
    param($ItemCount) 
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
    return @"
<html><head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: 'Consolas', monospace; background-color: #2D2D2E; color: #FFFFFF; padding: 20px; }
        .stats-bar { color: #00FFFF; font-size: 12px; border-bottom: 1px solid #555; padding-bottom: 10px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th { background-color: #483D8B; color: cyan; text-align: left; padding: 12px; border-bottom: 2px solid #555; cursor: pointer; user-select: none; }
        th:hover { background-color: #5A4FCF; }
        td { padding: 10px; border-bottom: 1px solid #444; }
        
        /* --- THE WIDTH LOCK --- */
        td:first-child, th:first-child { width: 110px; min-width: 110px; white-space: nowrap; }
        
        tr:hover { background-color: #3E3E42; }
        a { color: #FFFF00; text-decoration: none; }
        .paperclip { color: #00FFFF !important; text-decoration: none; font-size: 18px; cursor: pointer; }
        .paperclip:hover { color: #FFFFFF !important; }
    </style>
    <script>
        // 1. Run the saved sort as soon as the page loads
        window.onload = function() {
            const lastCol = localStorage.getItem('sortCol');
            const lastDir = localStorage.getItem('sortDir');
            if (lastCol !== null) {
                // Apply the sort without toggling (restoring existing state)
                sortTable(parseInt(lastCol), lastDir);
            }
        };

        function sortTable(n, forceDir = null) {
            var table = document.querySelector("table");
            var tbody = table.querySelector("tbody");
            if (!tbody) return;
            var rows = Array.from(tbody.rows);
            
            // 2. Determine direction: use saved direction or toggle current
            var currentDir = table.getAttribute("data-sort-dir-" + n);
            var dir = forceDir || (currentDir === "asc" ? "desc" : "asc");
            
            rows.sort(function(a, b) {
                var valA = a.cells[n].innerText.trim().toLowerCase();
                var valB = b.cells[n].innerText.trim().toLowerCase();
                
                // Handle date sorting specifically for column 0
                if (n === 0) {
                    return dir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
                }
                return dir === "asc" ? valA.localeCompare(valB) : valB.localeCompare(valA);
            });

            // 3. Save the state to the browser's memory
            localStorage.setItem('sortCol', n);
            localStorage.setItem('sortDir', dir);

            table.setAttribute("data-sort-dir-" + n, dir);
            rows.forEach(row => tbody.appendChild(row));
        }
    </script>
</head>
<body>
    <h1>$global:ScriptTitle1 | $global:Cvers | Build: $global:BuildDate | Master Index</h1>
    <div class="stats-bar">
        <span><b>Source:</b> $global:MasterPath</span>
        <span style='text-align:center;'><b>Destination:</b> $global:ArchivePath</span>
        <span style='text-align:right;'><b>Count:</b> $ItemCount | <b>Generated:</b> $timestamp</span>
    </div>
    <table>
        <thead>
            <tr>
                <th onclick="sortTable(0)">Date &#9662;</th>
                <th onclick="sortTable(1)">Subject &#9662;</th>
                <th onclick="sortTable(2)">From &#9662;</th>
                <th class="paperclip" onclick="sortTable(3)">&#128206;&#xFE0E;</th>
            </tr>
        </thead>
        <tbody>
"@
}

function Add-IndexRow { 
    param($Date, $Subject, $From, $MsgLink, $PdfLink) 
    # Using &#128206; (Paperclip) and &#xFE0E; (Text variation) safely
    $clip = if ($PdfLink) { "<a class='paperclip' href='$PdfLink' title='Open Attachment Folder'>&#128206;&#xFE0E;</a>" } else { "" } 
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
    if ([string]::IsNullOrWhiteSpace($global:ArchivePath)) {
        Update-Log "INDEX ERROR: Archive Path is empty." $true
        return
    }

    Update-Log "GENERATING MASTER INDEX..." $true
    $indexFile = Join-Path $global:ArchivePath "INDEX.html"
    # In Update-MasterIndex, update your Get-ChildItem line:
    $items = Get-ChildItem $global:ArchivePath -Exclude "INDEX.html", "*.meta", "*.dll", "*.bat", "*.ps1"

    $content = Get-HTMLHeader -ItemCount $items.Count

    $count = 0
    $total = $items.Count
    $Indextimer = [System.Diagnostics.Stopwatch]::StartNew()

    # --- THE ROW GENERATION LOOP ---
    foreach ($item in $items) {
        $count++
        Update-ArchivistProgress -Count $count -Total $total -Stopwatch $Indextimer -FileName $item.BaseName

        # 1. DEFAULTS
        $displaySubject = $item.BaseName
        $displayDate = $item.CreationTime.ToString("yyyy-MM-dd")
        $displaySender = "Archived Item"
        $link = if ($item.PSIsContainer) { "$($item.Name)/Message.html" } else { $item.Name }
        $pdfLink = ""

        # 2. DATA DISCOVERY: Look for Metadata (HTML or .meta sidecar)
        $rawData = $null
        if ($item.PSIsContainer) {
            $msgHtmlFile = Join-Path $item.FullName "Message.html"
            if (Test-Path $msgHtmlFile) { $rawData = Get-Content $msgHtmlFile -Raw }
            
            # Attachment Check for folders
            $hasAt = Get-ChildItem $item.FullName -Exclude "Message.html" | Select-Object -First 1
            if ($hasAt) { $pdfLink = "$($item.Name)" }
        }
        else {
            $metaFile = Join-Path $global:ArchivePath "$($item.BaseName).meta"
            if (Test-Path $metaFile) { $rawData = Get-Content $metaFile -Raw }
        }

        # 3. UNIFIED SCRAPER: If we found data, extract the 3 key fields
        if ($rawData) {
            if ($rawData -match "id='metadata-subject'>(.*?)</span>") { $displaySubject = $matches[1].Trim() }
            if ($rawData -match "id='metadata-from'>(.*?)</span>") { $displaySender = $matches[1].Trim() }
            if ($rawData -match "id='metadata-date'>(.*?)</span>") { $displayDate = $matches[1].Trim() }
        }

        # 4. ADD THE ROW
        $content += Add-IndexRow -Date $displayDate -Subject $displaySubject -From $displaySender -MsgLink $link -PdfLink $pdfLink
        [System.Windows.Forms.Application]::DoEvents()
    }

    # 5. SEAL FILE
    $content += "</tbody></table><div style='margin-top:20px;padding:10px;border-top:1px solid #555;color:cyan;font-family:Consolas;'>TOTAL: $($items.Count) Items</div></body></html>"
    Set-Content $indexFile -Value $content -Encoding utf8
    Update-Log "MASTER INDEX SEALED: $($items.Count) items." $true
    if ($global:chkAutoOpen.Checked) { Start-Process $indexFile }
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
    $masterReady = if ($global:rbSourcePST.Checked) { $null -ne $global:SelectedPstFolder } else { -not [string]::IsNullOrWhiteSpace($global:MasterPath) }
    
    if (-not $masterReady -or [string]::IsNullOrWhiteSpace($global:ArchivePath)) {
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

    # 1. Initialize State & UI
    $global:AbortArchive = $false
    $global:btnRun.Enabled = $false
    $global:btnAbort.Enabled = $true
    $global:btnAbort.Text = "ABORT"
    $global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed
    $global:Log.Clear()
    $global:ProgressBar.Value = 0

    Update-Log "INITIALIZING ENGINES..." $true
    $outlook = $null
    $word = $null

    try {
        # 2. Connect to Outlook COM
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $namespace.Logon("", "", $false, $false)

        # 3. Start Word Engine ONCE (Outside Loop)
        if ($global:rbDOCX.Checked) {
            Update-Log "STARTING WORD ENGINE..."
            try {
                $word = New-Object -ComObject Word.Application
                Start-Sleep -Seconds 2
                $word.DisplayAlerts = 0
                $word.Visible = $false
            }
            catch {
                Update-Log "FATAL: Could not start Word." $true
                return
            }
        }

        # 4. Gather Items
        if ($global:rbSourcePST.Checked) {
            if (-not $global:SelectedPstFolder) { Update-Log "ERROR: No PST Folder selected." $true; return }
            $itemsToProcess = $global:SelectedPstFolder.Items | Sort-Object Subject
        }
        else {
            # Clean the path string to ensure no hidden whitespace or trailing slashes
            $cleanPath = $global:MasterPath.Trim().TrimEnd('\')
            Update-Log "Searching in path: [$cleanPath]..." $true

            if (-not (Test-Path -Path $cleanPath)) {
                Update-Log "ERROR: Source path not found." $true; return
            }

            # Use -Path for the folder and -Filter for the extension
            $itemsToProcess = Get-ChildItem -Path $cleanPath -Filter "*.msg" | Sort-Object Name
        }

        # Count check with a 'null' safety
        $totalCount = if ($itemsToProcess) { $itemsToProcess.Count } else { 0 }
        Update-Log "Found ($totalCount) items to process." $true

        if ($totalCount -eq 0) {
            Update-Log "ERROR: No .msg files found in $cleanPath" $true
            return
        }

        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $count = 0

        # --- 5. THE MAIN LOOP ---
        foreach ($item in $itemsToProcess) {
            if ($global:AbortArchive) { break }

            if ($global:rbSourcePST.Checked) {
                $msg = $item
                $rawSubject = $msg.Subject
            }
            else {
                $msg = $outlook.CreateItemFromTemplate($item.FullName)
                Start-Sleep -Milliseconds 100 
                $rawSubject = $item.BaseName
            }

            # Safety Guard
            if ($msg.MessageClass -ne "IPM.Note") { 
                Update-Log "SKIPPING: $($rawSubject)"
                continue 
            }

            $count++

            # 2. CAPTURE THE UNIFIED META-DATA 
            $emailDate = $msg.ReceivedTime.ToString("yyyy-MM-dd")
            $emailFrom = if ($msg.SenderName) { $msg.SenderName } else { "Unknown Sender" }
            $fullSubject = $msg.Subject
            $safeName = ($rawSubject -replace '[\\\/\:\*\?\"\<\>\|\[\]]', '_').Trim(" -.")

            # 3. UPDATE THE UI
            Update-ArchivistProgress -Count $count -Total $totalCount -Stopwatch $timer -FileName $rawSubject

            # --- BRANCH A: HTML (PERFECTED) ---
            if ($global:rbHTML.Checked) {
                $itemFolder = Join-Path $global:ArchivePath $safeName
                $htmlPath = Join-Path $itemFolder "Message.html"

                if (-not (Test-Path $htmlPath)) {
                    if (-not (Test-Path $itemFolder)) { New-Item -ItemType Directory -Path $itemFolder -Force | Out-Null }
                    
                    $metadata = "<div style='display:none;'><span id='metadata-from'>$emailFrom</span><span id='metadata-date'>$emailDate</span><span id='metadata-subject'>$fullSubject</span></div>"
                    $attachLinks = "<div style='background:#333; padding:10px; border:1px solid #555; margin-bottom:15px; font-family:Consolas;'><b style='color:cyan;'>ATTACHMENTS:</b><br>"
                    
                    foreach ($at in $msg.Attachments) {
                        if ($at.FileName -notlike "image*") {
                            $safeAtName = ($at.FileName -replace '[\\\/\:\*\?\"\<\>\|\[\]]', '_').Trim(" -.")
                            $fullAtSavePath = Join-Path $itemFolder $safeAtName
                            try {
                                $at.SaveAsFile($fullAtSavePath)
                                $attachLinks += "<p style='margin:5px 0;'><a href='$safeAtName' style='color:cyan;'>[FILE] $safeAtName</a></p>"
                            }
                            catch { Update-Log "ATTACHMENT ERROR: $($at.FileName)" }
                        }
                        
                    } # ... inside Branch A: HTML, after the attachment loop ...
                    $attachLinks += "</div>"

                    # 1. THE HEADER (Simple text, no background, sits right above the email body)
                    $visibleHeader = "<div style='font-family:Consolas, monospace; color:black; margin-top:20px; margin-bottom:20px;'>"
                    $visibleHeader += "<b>FROM:</b> $emailFrom<br>"
                    $visibleHeader += "<b>DATE:</b> $emailDate<br>"
                    $visibleHeader += "<b>SUBJECT:</b> $fullSubject<br>"
                    $visibleHeader += "</div>"

                    # 2. THE FINAL ASSEMBLY
                    # We do NOT put a background style in <body> so the email body can breathe
                    $fullHtml = "<html><body>"
                    $fullHtml += $metadata        # Hidden
                    $fullHtml += $attachLinks     # Dark Banner (self-contained)
                    $fullHtml += $visibleHeader   # Cyan Header
                    $fullHtml += $($msg.HTMLBody) # Original Email Body (with its own styles)
                    $fullHtml += "</body></html>"

                    Set-Content -Path $htmlPath -Value $fullHtml -Encoding utf8
                    Update-Log "Archived (HTML): $safeName"
                }


            }

            # --- BRANCH B: DOCX (UNIFIED) ---
            elseif ($global:rbDOCX.Checked) {
                $docxPath = Join-Path $global:ArchivePath "$safeName.docx"
                # Inside the DOCX Branch
                if (-not (Test-Path $docxPath)) {
                    # Re-check that $word hasn't been disconnected
                    if ($null -eq $word) { throw "Word engine disconnected." }

                    $doc = $word.Documents.Add()
                    
                    $word.Selection.Font.Bold = $true
                    $word.Selection.TypeText("FROM: $emailFrom`nDATE: $emailDate`nSUBJECT: $fullSubject`n")
                    $word.Selection.TypeText("--------------------------------------------------`n")
                    $word.Selection.Font.Bold = $false

                    if ($msg.Attachments.Count -gt 0) {
                        $word.Selection.Font.Bold = $true
                        $word.Selection.TypeText("--- ATTACHED OBJECTS ---`n")
                        $word.Selection.Font.Bold = $false
                        foreach ($at in $msg.Attachments) {
                            if ($at.FileName -notlike "image*") {
                                $atSafe = ($at.FileName -replace '[\\\/\:\*\?\"\<\>\|\[\]]', '_').Trim(" -.")
                                $temp = Join-Path $env:TEMP $atSafe
                                try {
                                    $at.SaveAsFile($temp)
                                    $word.Selection.InlineShapes.AddOLEObject($null, $temp, $false, $true, $null, $null, $atSafe) | Out-Null
                                    $word.Selection.TypeParagraph()
                                    Remove-Item $temp -Force
                                }
                                catch { Update-Log "DOCX ATTACH ERROR: $($at.FileName)" }
                            }
                        }
                    }
                    # 2. Save and Close
                    $word.Selection.TypeText($msg.Body)
                    $doc.SaveAs([string]$docxPath)
                    $doc.Close()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                    # --- NEW: STAMP THE FILE WITH THE EMAIL DATE ---
                    $fileObj = Get-Item -Path $docxPath
                    $fileObj.CreationTime = $msg.ReceivedTime
                    $fileObj.LastWriteTime = $msg.ReceivedTime
                    # -----------------------------------------------
        
                    # --- NEW: CREATE A HIDDEN METADATA TAG FOR DOCX ---
                    $metaPath = Join-Path $global:ArchivePath "$safeName.meta"
                    # FIXED: Added the missing <span tags so the Scraper can find them
                    $metaContent = "<span id='metadata-from'>$emailFrom</span><span id='metadata-date'>$emailDate</span><span id='metadata-subject'>$fullSubject</span>"
    
                    Set-Content -Path $metaPath -Value $metaContent -Encoding utf8
    
                    # Optional: Make the .meta file invisible in Windows Explorer
                    (Get-Item $metaPath).Attributes = 'Hidden'

                    Update-Log "Archived (DOCX): $safeName"
                }
            }
        
            # --- BRANCH C: TXT (UNIFIED & TAGGED) ---
            elseif ($global:rbTXT.Checked) {
                $txtPath = Join-Path $global:ArchivePath "$safeName.txt"
            
                if (-not (Test-Path $txtPath)) {
                    # 1. Build a clean text header so the file is readable in Notepad
                    $txtHeader = "FROM: $emailFrom`r`nDATE: $emailDate`r`nSUBJECT: $fullSubject`r`n"
                    $txtHeader += "--------------------------------------------------`r`n`r`n"
                
                    $fullTxtBody = $txtHeader + $msg.Body
                
                    # 2. Save the TXT file
                    Set-Content -Path $txtPath -Value $fullTxtBody -Encoding utf8

                    # 3. Create the Hidden Sidecar for the Auditor
                    $metaPath = Join-Path $global:ArchivePath "$safeName.meta"
                    $metaText = "<span id='metadata-from'>$emailFrom</span><span id='metadata-date'>$emailDate</span><span id='metadata-subject'>$fullSubject</span>"
                    Set-Content -Path $metaPath -Value $metaText -Encoding utf8
                    (Get-Item $metaPath).Attributes = 'Hidden'

                    # 4. Stamp the file with the actual Email Date
                    $fileObj = Get-Item -Path $txtPath
                    $fileObj.CreationTime = $msg.ReceivedTime
                    $fileObj.LastWriteTime = $msg.ReceivedTime

                    Update-Log "Archived (TXT): $safeName"
                }
            }
        } # End Foreach Loop

        if ($global:chkIndex.Checked -and -not $global:AbortArchive) { Update-MasterIndex }

    }
    catch {
        Update-Log "CRITICAL ERROR: $_" $true
    }
    finally {
        if ($word) { $word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
        if ($outlook) { $outlook.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
        
        $global:btnRun.Enabled = $true
        $global:btnAbort.Enabled = $false
        $global:btnAbort.Text = "ABORT"
        $global:btnAbort.BackColor = [System.Drawing.Color]::DarkRed

        # FIX: Only stop the timer if it was actually initialized
        if ($null -ne $timer) { $timer.Stop() }

        $global:lblAction.Text = "System Ready."
    }
}

# 4.5 Check4Update Set your current version (Update this every time you release!)
function Test-Update {
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send("1.1.1.1", 1000)
        if ($reply.Status -ne "Success") { return }
    } catch { return }

    try {
        $apiUrl = "https://api.github.com/repos/$($global:RepoOwner)/$($global:RepoName)/releases/latest"
        $userAgent = "THE-ARCHIVIST-Script"
        
        $release = Invoke-RestMethod -Uri $apiUrl -Method Get -UserAgent $userAgent -ContentType "application/json" -ErrorAction Stop 
        
        if ($null -eq $release) { return } 
        
        $rawTagName = $release.tag_name
        
        $latestStr = $rawTagName -replace '[^0-9.]', '' 
        $currentStr = $global:Cvers -replace '[^0-9.]', '' 
        
        if ([string]::IsNullOrWhiteSpace($latestStr) -or [string]::IsNullOrWhiteSpace($currentStr)) { 
            return 
        } 
        
        if ([version]$latestStr -gt [version]$currentStr) { 
            # Turn the GitHub link CYAN to indicate an update is available
            $global:lblGitHub.LinkColor = [System.Drawing.Color]::Orange
            
            Add-Type -AssemblyName System.Windows.Forms 
            Add-Type -AssemblyName System.Drawing
            
            $downloadUrl = "https://github.com/$($global:RepoOwner)/$($global:RepoName)/releases/latest"
            
            # Create a custom form with clickable link
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Update Available"
            $form.Width = 450
            $form.Height = 210
            $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $form.MaximizeBox = $false
            $form.MinimizeBox = $false
            
            # Add title label
            $titleLabel = New-Object System.Windows.Forms.Label
            $titleLabel.Text = "A new version of $($global:ScriptTitle1) ($rawTagName) is available!"
            $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
            $titleLabel.Width = 410
            $titleLabel.Height = 30
            $titleLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $form.Controls.Add($titleLabel)
            
            # Add message with clickable link
            $messagePanel = New-Object System.Windows.Forms.Panel
            $messagePanel.Location = New-Object System.Drawing.Point(20, 60)
            $messagePanel.Width = 410
            $messagePanel.Height = 50
            $messagePanel.AutoSize = $true
            
            # "To Download: " text
            $downloadLabel = New-Object System.Windows.Forms.Label
            $downloadLabel.Text = "To Download: "
            $downloadLabel.Location = New-Object System.Drawing.Point(0, 0)
            $downloadLabel.Width = 100
            $downloadLabel.Height = 25
            $downloadLabel.AutoSize = $true
            $messagePanel.Controls.Add($downloadLabel)
            
            # "Click Here" link (keep dark blue)
            $clickLink = New-Object System.Windows.Forms.LinkLabel
            $clickLink.Text = "Click Here"
            $clickLink.Location = New-Object System.Drawing.Point(90, 0)
            $clickLink.Width = 100
            $clickLink.Height = 25
            $clickLink.AutoSize = $true
            $clickLink.LinkColor = [System.Drawing.Color]::Blue
            $clickLink.add_LinkClicked({
                [System.Diagnostics.Process]::Start($downloadUrl)
            })
            $messagePanel.Controls.Add($clickLink)
            
            $form.Controls.Add($messagePanel)
            
            # Add OK button
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(185, 115)
            $okButton.Width = 80
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Controls.Add($okButton)
            $form.AcceptButton = $okButton
            
            # Add "Check for Updates" link at bottom left (blue reminder)
            $updateLink = New-Object System.Windows.Forms.LinkLabel
            $updateLink.Text = "Check for Updates"
            $updateLink.Location = New-Object System.Drawing.Point(20, 175)
            $updateLink.Width = 150
            $updateLink.Height = 25
            $updateLink.AutoSize = $true
            $updateLink.LinkColor = [System.Drawing.Color]::Blue
            $updateLink.add_LinkClicked({
                [System.Diagnostics.Process]::Start($downloadUrl)
            })
            $form.Controls.Add($updateLink)
            
            $form.ShowDialog() | Out-Null
        }
        else {
            # User is on the latest version - turn the GitHub link back to GRAY
            $global:lblGitHub.LinkColor = [System.Drawing.Color]::Gray
        }
    } 
    catch { 
        Write-Debug "Update check failed: $($_.Exception.Message)" 
    } 
}

# ==============================================================================
# SECTION 5: EVENT WIRING (Connecting the Dots)
# ==============================================================================

# --- Section 5: Wiring the btnMaster (Corrected Nesting) ---
$global:btnMaster.Add_Click({
        if ($global:rbSourcePST.Checked) {
            # --- MODE: PST ---
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
            $FileBrowser.Filter = "Outlook Data File (*.pst) | *.pst"
            if ($FileBrowser.ShowDialog($global:Form) -eq [Windows.Forms.DialogResult]::OK) {
                try {
                    $outlook = New-Object -ComObject Outlook.Application
                    $namespace = $outlook.GetNamespace("MAPI")
                    $namespace.AddStore($FileBrowser.FileName)
                    $selectedFolder = $namespace.PickFolder()
                
                    if ($selectedFolder) {
                        $global:MasterPath = $selectedFolder.FolderPath
                        $global:SelectedPstFolder = $selectedFolder
                    
                        # Apply your preferred feedback style
                        $global:btnMaster.Text = "PST: " + $selectedFolder.Name
                        $global:btnMaster.BackColor = [System.Drawing.Color]::DarkGreen
                        Update-Log "MASTER PST SOURCE SET: $($selectedFolder.FolderPath)"
                    
                        # Logic Gate: Check if Archive is also set
                        if ($global:MasterPath -ne "" -and $global:ArchivePath -ne "") {
                            $global:btnRun.Enabled = $true
                            $global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
                            Update-Log "System Ready. Select options and press START."
                        }
                    }
                }
                catch { Update-Log "Outlook Connection Error." $true }
            }
        }
        else {
            # --- MODE: FOLDER (Uses your logic exactly) ---
            $fb = New-Object Windows.Forms.FolderBrowserDialog
            $fb.ShowNewFolderButton = $true
            if ($fb.ShowDialog($global:Form) -eq [Windows.Forms.DialogResult]::OK) {
                $global:MasterPath = $fb.SelectedPath
                $global:btnMaster.Text = "MASTER: " + (Split-Path $global:MasterPath -Leaf)
                $global:btnMaster.BackColor = [System.Drawing.Color]::DarkGreen
                Update-Log "MASTER SOURCE SET: $global:MasterPath"
            
                if ($global:MasterPath -ne "" -and $global:ArchivePath -ne "") {
                    $global:btnRun.Enabled = $true
                    $global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
                    Update-Log "System Ready. Select options and press START."
                }
            }
        }
    })


$global:btnTarget.Add_Click({ Set-Path -Type "Target" })
# Wiring the Rebuild button
$global:btnRebuild.Add_Click({ Update-MasterIndex })

# Update the Start button to call the Indexer at the end
$global:btnRun.Add_Click({ 
        Invoke-ArchivalProcess
    })

$global:btnReset.Add_Click({
        # 1. Clear Paths
        $global:MasterPath = ""
        $global:ArchivePath = ""

        # 2. Reset Text
        $global:btnMaster.Text = "Sel MASTER (.msg source)"
        $global:btnTarget.Text = "Sel ARCHIVE (destination)"

        # 3. Defaults & Dashboard
        $global:rbHTML.Checked = $true
        $global:rbDOCX.Checked = $false
        $global:chkExtract.Enabled = $true
        $global:chkExtract.Checked = $true
        $global:btnRun.Enabled = $false
        $global:btnRun.BackColor = [System.Drawing.Color]::DarkSlateBlue
        $global:ProgressBar.Value = 0
        $global:lblAction.Text = "Ready to Archive..."
        Update-Log "--- SYSTEM RESET TO DEFAULTS ---" $true

        # 4. THE UI FIX: Set specific dark colors for Master & Target
        $buttons = @($global:btnMaster, $global:btnTarget)
        # 4. THE UI FIX: Set specific dark colors and restore Cyan Borders
        $buttons = @($global:btnMaster, $global:btnTarget)
        foreach ($btn in $buttons) {
            $btn.BackColor = [System.Drawing.Color]::FromArgb(255, 64, 64, 70)
            $btn.ForeColor = [System.Drawing.Color]::Cyan
        
            # RESTORE THE CYAN BORDER
            $btn.FlatAppearance.BorderSize = 1
            $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::Cyan 
        
            # KEEP THE HOVER DARK
            $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 80, 80, 85)
            $btn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkSlateBlue
        }
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
            Update-Log "DOCX Mode: Extraction disabled (Attachments will be EMBEDDED)."
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
        $global:btnAbort, 
        $global:btnRebuild, 
        $global:ProgressBar, 
        $global:lblAction, 
        $global:Log, 
        $global:lblGitHub,
        $global:lblSupportText, 
        $global:lblSupportLink
    ))
# This triggers the update check as soon as the app starts
Test-Update
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
