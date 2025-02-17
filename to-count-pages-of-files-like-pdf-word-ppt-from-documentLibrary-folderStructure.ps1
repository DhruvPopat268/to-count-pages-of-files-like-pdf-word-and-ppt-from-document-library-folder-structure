# Connect to SharePoint Online
Connect-PnPOnline -Url "https://futurrizoninterns.sharepoint.com/sites/lookUpDataTesting" -UseWebLogin

# Load PDFSharp Library (Ensure the path is correct)
Add-Type -Path "C:\Users\91915\Downloads\pdfsharpcore.1.3.65\lib\net7.0\PdfSharpCore.dll"

# Define the destination folder
$destinationFolder = "C:\Users\91915\Downloads\nuget Project"

# Ensure the destination folder exists, create it if not
if (!(Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder -Force
}

# Function to count pages in PDF files
function Get-PDFPageCount {
    param ([string]$pdfPath)
    try {
        $document = [PdfSharpCore.Pdf.IO.PdfReader]::Open($pdfPath, [PdfSharpCore.Pdf.IO.PdfDocumentOpenMode]::Import)
        $pageCount = $document.PageCount
        $document.Close()
        Write-Host "PDF: $pdfPath - Pages: $pageCount"
        return $pageCount
    } catch {
        Write-Host "Error processing PDF file: $pdfPath - $_"
        return $null
    }
}

# Function to count pages in Word documents
function Get-WordPageCount {
    param ([string]$FilePath)
    try {
        Write-Host "Processing Word file: $FilePath"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false  # Ensure Word is hidden
        $doc = $word.Documents.Open($FilePath, [Type]::Missing, $true)
        $pageCount = $doc.ComputeStatistics(2)  # wdStatisticPages = 2
        $doc.Close($false)
        $word.Quit()
        Write-Host "Word: $FilePath - Pages: $pageCount"
        return $pageCount
    } catch {
        Write-Host "Error reading Word file: $FilePath - $_"
        return $null
    } finally {
        Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
    }
}

# Function to count slides in PowerPoint files
function Get-PPTSlideCount {
    param ([string]$FilePath)
    try {
        Write-Host "Processing PowerPoint file: $FilePath"
        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = 2  # Use '2' for hidden mode instead of $false
        $presentation = $ppt.Presentations.Open($FilePath, $false, $false, $false)
        $slideCount = $presentation.Slides.Count
        $presentation.Close()
        $ppt.Quit()
        Write-Host "PowerPoint: $FilePath - Slides: $slideCount"
        return $slideCount
    } catch {
        Write-Host "Error reading PowerPoint file: $FilePath - $_"
        return $null
    } finally {
        Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
    }
}

# Set the correct SharePoint Document Library name
$libraryName = "Document Management Library 2"

# Ensure the PageCount column exists
$fields = Get-PnPField -List $libraryName | Select InternalName, Title
if ($fields.InternalName -notcontains "PageCount") {
    Write-Host "Error: 'PageCount' column is missing. Ensure it exists in the SharePoint list."
    exit
}

# Retrieve all files, including nested folders, from the SharePoint Document Library
$files = Get-PnPListItem -List $libraryName -Fields "FileRef", "ID" | Where-Object { $_["FileRef"] -match "\.(pdf|docx|pptx)$" -and $_["FileRef"] -notmatch "/Forms/" }

Write-Host "Total Files Found: $($files.Count)"

# Process each file and update PageCount
foreach ($file in $files) {
    $fileUrl = $file["FileRef"]
    $fileName = [System.IO.Path]::GetFileName($fileUrl)
    $localPath = "$destinationFolder\$fileName"

    # Download the file to the specified folder
    Write-Host "Downloading: $fileUrl to $localPath"
    Get-PnPFile -Url $fileUrl -Path $destinationFolder -FileName $fileName -AsFile -Force

    # Initialize page/slide count
    $PageCount = $null

    # Determine the file type and get the page/slide count
    if ($fileName -match "\.pdf$") {
        $PageCount = Get-PDFPageCount -pdfPath $localPath
    } elseif ($fileName -match "\.docx$") {
        $PageCount = Get-WordPageCount -FilePath $localPath
    } elseif ($fileName -match "\.pptx$") {
        $PageCount = Get-PPTSlideCount -FilePath $localPath
    }

    # Update SharePoint column if page count is retrieved
    if ($PageCount -ne $null) {
        Write-Host "Updating SharePoint: $fileUrl with $PageCount pages/slides."
        Set-PnPListItem -List $libraryName -Identity $file.Id -Values @{"PageCount" = $PageCount}
    } else {
        Write-Host "Skipping update for $fileUrl due to missing page/slide count."
    }

    # Clean up the local file (optional)
    Remove-Item -Path $localPath -Force
}

Write-Host "Processing Completed!"
