# MSG-to-DOCX-Batch-Converter
MSG to DOCX Batch Converter with Embedded PDF Attachments
# MSG to DOCX Batch Converter with Embedded PDF Attachments

A lightweight, free automation tool to batch convert Microsoft Outlook `.msg` files into modern Word `.docx` documents while physically embedding any attached PDF files directly into the page. 

This project was born out of frustration with finding only paid, third-party software to extract and convert old archived `.msg` files. This solution is 100% free and runs locally on your machine.

## 🌟 Features
* **Zero Cost:** No paid licenses or third-party paid tools required.
* **Preserves Attachments:** PDF attachments are embedded as double-clickable objects inside the Word document.
* **Organized Output:** Automatically creates a secondary folder to store your converted documents without cluttering your original files.
* **Archival Friendly:** Converts proprietary binary `.msg` files into open-standard `.docx` files, making them readable by LibreOffice or OpenOffice.

## 📋 Prerequisites
Because this script uses COM object automation to tap into native desktop applications, you must have the following installed on your Windows machine:
* **Microsoft Outlook** (Desktop application)
* **Microsoft Word** (Desktop application)
* **PowerShell** (Built into Windows)

*Note: Outlook and Word are only required on the host machine performing the conversion. The finished documents can be opened on any computer using free software like LibreOffice!*

## 🚀 How to Use

### 1. Repository Setup
Download files from this repository, (the `.ps1` script and the `.bat` file) and place them into the folder where your `.msg` files are stored (e.g., `C:\MyEmails`).

### 2. Path Verification
Open the `ConvertEmails.ps1` file in Notepad and ensure the folder path at the top matches your actual folder:
powershell
$folderPath = "C:\MyEmails"  is the DEFAULT PATH.  Feel free to change this if you choose.

### 3. Execute
1.	Ensure the Microsoft Outlook desktop app is completely closed to prevent security memory conflicts.
2.	Double-click the Launch.bat file.
3.	Watch the terminal convert your files! Your new files will be safely deposited in a new Converted_Files sub-directory.

## ⚖️ License
This project is open-source and free to use, modify, and share. Pay it forward!

## ☕ Support the Project

If "ConvertEmails.ps1" saved you some time or made your conversion easier, feel free to help keep the gears turning!

* [**Support via Ko-fi**](https://ko-fi.com/kofisupporter19535)
* [**Support via Buy Me a Coffee**](https://www.buymeacoffee.com/jj1eckhardt)

