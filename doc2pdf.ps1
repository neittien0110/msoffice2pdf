# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
# Thanks to MFT, takabanana, ComFreek
#
# Adapted from https://gist.github.com/allenyllee/5d7c4a16ae0e33375e4a6d25acaeeda2
# Thank to mp4096, the author of the script ppt2pdf which I customized to this script
##
## about_Execution_Policies | Microsoft Docs
## https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1&viewFallbackFrom=powershell-Microsoft.PowerShell.Core
## Beginning in Windows PowerShell 3.0, you can use the Stream parameter of the Get-Item cmdlet to detect files that are blocked because they were downloaded from the Internet, and you can use the Unblock-File cmdlet to unblock the scripts so that you can run them in Windows PowerShell.
##
## To execute without error, Run below command:
## ```
## Unblock-File .\doc2pdf.ps1; powershell -ExecutionPolicy RemoteSigned -File .\doc2pdf.ps1
## ```
## or do like the image: https://4sysops.com/wp-content/uploads/2015/01/Unblock-in-File-Explorer.png

## MSDN Search
## https://social.msdn.microsoft.com/Search/en-US

##
## $MyInvocation.MyCommand.Path – Noam's scripting blog
## https://scriptingblog.com/tag/myinvocation-mycommand-path/
##
## Get commandline params
Param(
    ## File ppt? must contains this word
    [Parameter(Mandatory=$false)]
    [string]
    $NameCondition,

    ## Default folder
    [Parameter(Mandatory=$false)]
    [string]
    $Folder
)



# Get invocation path
$scriptpath = $MyInvocation.MyCommand.Path
if ($Folder) {
    $curr_path = $Folder; 
} else {
    $curr_path = Split-Path $scriptpath;
}

Write-Host "Scanning in folder "  $curr_path "to convert files contains "$NameCondition -BackgroundColor Magenta

##
## New-Object - PowerShell - SS64.com
## https://ss64.com/ps/new-object.html
##
# Create a Word object
#$doc_app = New-Object -ComObject Word.Application
#
# or
#
##
## Using Interop Assemblies | Richard Siddaway's Blog
## https://richardspowershellblog.wordpress.com/2007/05/29/using-interop-assemblies/
##
# Load Word Interop Assembly
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Word") > $null
[Reflection.Assembly]::LoadWithPartialname("Office") > $null # need this or Word might not close
$doc_app = New-Object "Microsoft.Office.Interop.Word.ApplicationClass" 

# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.doc? | ForEach-Object {
    ##
    ## powershell - What is the difference between $_. and $_ - Stack Overflow
    ## https://stackoverflow.com/questions/35209737/what-is-the-difference-between-and
    ## 
    ## Get-ChildItem -Path C:\Windows | ForEach-Object {
    ##     $_  # this references the entire object returned.
    ## 
    ##     $_.FullName  # this refers specifically to the FullName property
    ## }
    ##

    ## Matching the Name filter
    if (-Not $_.Name.Contains($NameCondition)) {
        Write-Host ( "Skip " + $_.Name)
        return 
    }
        
    ##
    ## powershell - Which should I use: "Write-Host", "Write-Output", or "[console]::WriteLine"? - Stack Overflow
    ## https://stackoverflow.com/questions/8755497/which-should-i-use-write-host-write-output-or-consolewriteline
    ##
    ## Write-Output should be used when you want to send data on in the pipe line, but not necessarily want to display it on screen. 
    ## The pipeline will eventually write it to out-default if nothing else uses it first. 
    ## Write-Host should be used when you want to do the opposite.
    ##
    # print to screen...
    Write-Host "Processing" $_.FullName "..."
    
    # Open it in Word
    $document = $doc_app.Documents.Open($_.FullName)
    
    ##
    ## .net - powershell - extract file name and extension - Stack Overflow
    ## https://stackoverflow.com/questions/9788492/powershell-extract-file-name-and-extension
    ##
    ## If the file is coming off the disk and as others have stated, use the BaseName and Extension properties
    ##
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source Word file, replace $curr_path with $_.DirectoryName
    #$pdf_filename = "$($curr_path)\$($_.BaseName).pdf"
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    
    ##
    ## Presentation.SaveAs Method (PowerPoint)
    ## https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-saveas-method-powerpoint
    ##
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    #$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    #$document.SaveAs($pdf_filename, $opt)

   
    # String	The path for the export.
    $exportPath = $pdf_filename
    
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdexportformat?view=word-pia
    # wdExportFormatPDF	17	Export document into PDF format.
    # wdExportFormatXPS	18	Export document into XML Paper Specification (XPS) format.
    $fixedFormatType = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
    
     # https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat2
    $openAfter = $false

    # https://docs.microsoft.com/en-us/office/vba/api/word.wdexportoptimizefor
    # wdExportOptimizeForOnScreen	Export for screen, which is a lower quality and results in a smaller file size.
    # wdExportOptimizeForPrint	    Export for print, which is higher quality and results in a larger file size.
    $intent = [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint
    
    # https://docs.microsoft.com/en-us/office/vba/api/word.wdexportrange
    # Name	                Value	Description
    # wdExportAllDocument	0	    Exports the entire document.
    # wdExportCurrentPage	2	    Exports the current page.
    # wdExportFromTo	    3	    Exports the contents of a range using the starting and ending positions.
    # wdExportSelection	    1	    Exports the contents of the current selection
    $printRange = [Microsoft.Office.Interop.Word.wdexportrange]::wdExportAllDocument
    
    #$From = $document.PrintOptions.Ranges.Add(1, $document.Slides.Count)

    $item = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent

    # Boolean	Whether the document properties should also be exported. The default is False.
    $includeDocProperties = $false

    # https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat2
    # Boolean	Whether the IRM settings should also be exported. The default is True.
    $keepIRMSettings = $true
    

    # https://docs.microsoft.com/en-us/office/vba/api/word.wdexportcreatebookmarks
    $CreateBookmarks = [Microsoft.Office.Interop.Word.wdexportcreatebookmarks]::wdExportCreateWordBookmarks;

    # Boolean	Whether to include document structure tags to improve document accessibility. The default is True.
    $docStructureTags = $true
    
    # Boolean	Whether to include a bitmap of the text. The default is True.
    $bitmapMissingFonts = $true
    
    # Boolean	Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False.
    $useISO19005_1 = $false

    $OptimizeForImageQuality = $false
    
    # Boolean	Whether the resulting document should include associated pen marks.
    $FixedFormatExtClassPtr = $null
    
    ##
    ## Publishes as PDF or XPS.
    ##
    ## vba - difference between ExportAsFixedFormat2 and ExportAsFixedFormat? - Stack Overflow
    ## https://stackoverflow.com/questions/37585025/difference-between-exportasfixedformat2-and-exportasfixedformat
    ##
    
    # Document.ExportAsFixedFormat2 method (Word)
    # https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat2
    $document.ExportAsFixedFormat2($exportPath, $fixedFormatType, $openAfter, $intent, $printRange, 1, 2,  $item, $includeDocProperties, $keepIRMSettings, $CreateBookmarks, $docStructureTags, $bitmapMissingFonts, $useISO19005_1, $OptimizeForImageQuality)
    
    # Document.ExportAsFixedFormat method (Word)
    # https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
    #$document.ExportAsFixedFormat(..)
    
    # Close Word file
    $document.Close()
}
# Exit and release the Word object
$doc_app.Quit()

# Make sure references to COM objects are released, otherwise Word might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();

