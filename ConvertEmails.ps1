##Salvation is free, but it isn't cheap.  

# 1. Define the folder containing your 250 .msg files
$folderPath = "C:\MyEmails" 
$outputFolder = Join-Path $folderPath "Converted_Files"

# Create the secondary directory if it doesn't exist
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# 2. Open the background engines
$outlook = New-Object -ComObject Outlook.Application
$word = New-Object -ComObject Word.Application
$word.Visible = $false # Keep Word hidden in the background

# 3. Get all .msg files in that folder
$files = Get-ChildItem -Path $folderPath -Filter *.msg

foreach ($file in $files) {
    try {
        # Open the message using Outlook's template engine
        $msg = $outlook.CreateItemFromTemplate($file.FullName)
        
        # Create a brand new Word document
        $doc = $word.Documents.Add()
        $selection = $word.Selection
        
        # Write the email text into the Word document
        $selection.TypeText($msg.Body)
        $selection.TypeParagraph()
        
        # 4. Check for attachments and embed PDFs
        if ($msg.Attachments.Count -gt 0) {
            $selection.TypeParagraph()
            $selection.TypeText("--- ATTACHED PDF OBJECTS ---")
            $selection.TypeParagraph()
            
            foreach ($attachment in $msg.Attachments) {
                if ($attachment.FileName -like "*.pdf") {
                    
                    # Temporarily save the PDF so Word can grab it
                    $tempPdfPath = Join-Path $folderPath "temp_attachment.pdf"
                    $attachment.SaveAsFile($tempPdfPath)
                    
                    # Embed the PDF file as a clickable object into the Word doc
                    $selection.InlineShapes.AddOLEObject($null, $tempPdfPath, $false, $true, $null, $null, $attachment.FileName) | Out-Null
                    $selection.TypeParagraph()
                    
                    # Delete the temporary PDF file
                    Remove-Item $tempPdfPath -Force
                }
            }
        }
        
        # Define the new .docx file path in the subfolder
        $newFileName = $file.BaseName + ".docx"
        $newFilePath = Join-Path $outputFolder $newFileName
        
        # FIX APPLIED HERE: Forcing the path to be a pure string
        $doc.SaveAs([string]$newFilePath)
        $doc.Close()
        
        Write-Host "Successfully converted with PDFs: $($file.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error converting $($file.Name): $_" -ForegroundColor Red
    }
}

# 5. Clean up and close the background processes
$outlook.Quit()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "`nProcess complete! Press any key to close."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
