# to-count-pages-of-files-like-pdf-word-and-ppt-from-document-library-folder-structure

# SharePoint File Downloader and Page Counter

## Overview
This PowerShell script automates the process of downloading files from a **SharePoint Online Document Library** and counting pages in **PDF, Word, and PowerPoint files**. The script updates the page/slide count in a SharePoint metadata column named **PageCount**.

## Features
- ✅ **Automated SharePoint File Download**: Fetches PDF, DOCX, and PPTX files from SharePoint.
- ✅ **Page Counting**: Uses `PdfSharpCore`, Microsoft Word, and PowerPoint COM objects to count pages/slides.
- ✅ **Metadata Update**: Stores the page count in a SharePoint column.
- ✅ **Supports Nested Folders**: Retrieves all documents from subfolders.
- ✅ **Configurable Download Path**: Saves files to `C:\Users\91915\Downloads\nuget Project`.

## Prerequisites
Before running the script, ensure you have the following:

- **PowerShell 5.1+**
- **PnP PowerShell Module** (`Install-Module PnP.PowerShell`)
- **Microsoft Word & PowerPoint (for COM automation)**
- **PDFSharpCore Library** (`PdfSharpCore.dll` must be in the correct path)
- **Access to SharePoint Online**

## Installation
1. **Clone the Repository**
   ```sh
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo
   ```

2. **Install Required Modules**
   ```powershell
   Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
   ```

3. **Ensure the `PdfSharpCore.dll` Library is Available**
   - Download and place `PdfSharpCore.dll` in `C:\Users\91915\Downloads\pdfsharpcore.1.3.65\lib\net7.0\`

4. **Modify the Script (if needed)**
   - Update `Connect-PnPOnline` with your SharePoint site URL.
   - Change `$destinationFolder` to your preferred download location.

## Usage
1. **Run the Script**
   ```powershell
   .\DownloadAndProcessFiles.ps1
   ```

2. **Expected Output**
   - Files are downloaded to `C:\Users\91915\Downloads\nuget Project`
   - Page/slide counts are updated in SharePoint metadata.
   - The script prints logs in the PowerShell console.

## Script Breakdown
### Connect to SharePoint
```powershell
Connect-PnPOnline -Url "https://futurrizoninterns.sharepoint.com/sites/lookUpDataTesting" -UseWebLogin
```
### Download Files
```powershell
Get-PnPFile -Url $fileUrl -Path $destinationFolder -FileName $fileName -AsFile -Force
```
### Count Pages in PDFs
```powershell
$document = [PdfSharpCore.Pdf.IO.PdfReader]::Open($pdfPath, [PdfSharpCore.Pdf.IO.PdfDocumentOpenMode]::Import)
$pageCount = $document.PageCount
```
### Count Pages in Word Documents
```powershell
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($FilePath, [Type]::Missing, $true)
$pageCount = $doc.ComputeStatistics(2)
```
### Count Slides in PowerPoint Files
```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $ppt.Presentations.Open($FilePath, $false, $false, $false)
$slideCount = $presentation.Slides.Count
```
### Update SharePoint Metadata
```powershell
Set-PnPListItem -List $libraryName -Identity $file.Id -Values @{"PageCount" = $PageCount}
```

## Troubleshooting
- **PnP PowerShell Not Installed?** Run: `Install-Module -Name PnP.PowerShell`
- **Cannot Load PDFSharpCore.dll?** Check the correct path and permissions.
- **COM Automation Errors?** Ensure Microsoft Word & PowerPoint are installed.

## License
This project is licensed under the **MIT License**.

## Contributing
Pull requests are welcome! For major changes, open an issue first.

## Contact
- **Author**: Your Name
- **GitHub**: [yourusername](https://github.com/yourusername)

---
**Keywords:** PowerShell, SharePoint Online, Office 365, PnP PowerShell, File Processing, PDF Page Count, Word Page Count, PowerPoint Slide Count, SharePoint Automation, Metadata Update

