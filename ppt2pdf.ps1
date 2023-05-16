# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
# Thanks to MFT, takabanana, ComFreek
#
##
## about_Execution_Policies | Microsoft Docs
## https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1&viewFallbackFrom=powershell-Microsoft.PowerShell.Core
## Beginning in Windows PowerShell 3.0, you can use the Stream parameter of the Get-Item cmdlet to detect files that are blocked because they were downloaded from the Internet, and you can use the Unblock-File cmdlet to unblock the scripts so that you can run them in Windows PowerShell.
##
## To execute without error, Run below command:
## ```
## Unblock-File .\ppt2pdf.ps1; powershell -ExecutionPolicy RemoteSigned -File .\ppt2pdf.ps1
## ```

## MSDN Search
## https://social.msdn.microsoft.com/Search/en-US

##
## $MyInvocation.MyCommand.Path – Noam's scripting blog
## https://scriptingblog.com/tag/myinvocation-mycommand-path/
##
# Get invocation path
$scriptpath = $MyInvocation.MyCommand.Path
$curr_path = Split-Path $scriptpath


##
## New-Object - PowerShell - SS64.com
## https://ss64.com/ps/new-object.html
##
# Create a PowerPoint object
#$ppt_app = New-Object -ComObject PowerPoint.Application
#
# or
#
##
## Using Interop Assemblies | Richard Siddaway's Blog
## https://richardspowershellblog.wordpress.com/2007/05/29/using-interop-assemblies/
##
# Load Powerpoint Interop Assembly
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
[Reflection.Assembly]::LoadWithPartialname("Office") > $null # need this or powerpoint might not close
$ppt_app = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
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
    
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    
    ##
    ## .net - powershell - extract file name and extension - Stack Overflow
    ## https://stackoverflow.com/questions/9788492/powershell-extract-file-name-and-extension
    ##
    ## If the file is coming off the disk and as others have stated, use the BaseName and Extension properties
    ##
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    #$pdf_filename = "$($curr_path)\$($_.BaseName).pdf"
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    
    ##
    ## Presentation.SaveAs Method (PowerPoint)
    ## https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-saveas-method-powerpoint
    ##
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    #$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    #$document.SaveAs($pdf_filename, $opt)
    
    
    ##
    ## Export PDF with pen markups
    ## Based on this:
    ## https://github.com/netoffice/NetOffice_Import_from_SVN/blob/d44862542576ad1f2b0c7e42421aa42c8a3053c7/Source/PowerPoint/DispatchInterfaces/_Presentation.cs#L2943-L2969
    ## this:
    ## command line - How can I automatically convert PowerPoint to PDF? - Super User
    ## https://superuser.com/questions/641471/how-can-i-automatically-convert-powerpoint-to-pdf
    ## And this:
    ## Powershell script to export Powerpoint Presentations to pdf using the Powerpoint COM API
    ## https://gist.github.com/ap0llo/05cef76e3c4040ee924c4cfeef3f0b40
    ##
    
    # String	The path for the export.
    $exportPath = $pdf_filename
    
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppfixedformattype-enumeration-powerpoint
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.powerpoint.ppfixedformattype(v=office.14).aspx
    # ppFixedFormatTypeXPS	XPS format
    # ppFixedFormatTypePDF	PDF format
    $fixedFormatType = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF
    
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppfixedformatintent-enumeration-powerpoint
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.powerpoint.ppfixedformatintent(v=office.14).aspx
    # ppFixedFormatIntentScreen	Intent is to view exported file on screen.
    # ppFixedFormatIntentPrint	Intent is to print exported file.
    $intent = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentScreen
    
    # https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/msotristate-enumeration-office
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.core.msotristate.aspx
    # msoTrue	True.
    # msoFalse	False.
    # msoCTrue	Not supported.
    # msoTriStateToggle	Not supported.
    # msoTriStateMixed	Not supported.
    $frameSlides = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppprinthandoutorder-enumeration-powerpoint
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.powerpoint.ppprinthandoutorder(v=office.14).aspx
    # ppPrintHandoutVerticalFirst	Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it.
    #								If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide to the left of it.
    # ppPrintHandoutHorizontalFirst	Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it.
    #								If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide to the left of it.
    $handoutOrder = [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutVerticalFirst
    
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppprintoutputtype-enumeration-powerpoint
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.powerpoint.ppprintoutputtype(v=office.14).aspx
    # ppPrintOutputSlides	Slides
    # ppPrintOutputTwoSlideHandouts	Two Slide Handouts
    # ppPrintOutputThreeSlideHandouts	Three Slide Handouts
    # ppPrintOutputSixSlideHandouts	Six Slide Handouts
    # ppPrintOutputNotesPages	Notes Pages
    # ppPrintOutputOutline	Outline
    # ppPrintOutputBuildSlides	Build Slides
    # ppPrintOutputFourSlideHandouts	Four Slide Handouts
    # ppPrintOutputNineSlideHandouts	Nine Slide Handouts
    # ppPrintOutputOneSlideHandouts	Single Slide Handouts
    $outputType = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSlides
    
    # https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/msotristate-enumeration-office
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.core.msotristate.aspx
    # msoTrue	True.
    # msoFalse	False.
    # msoCTrue	Not supported.
    # msoTriStateToggle	Not supported.
    # msoTriStateMixed	Not supported.
    $printHiddenSlides = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    # Slides.Count 屬性 (PowerPoint)
    # https://msdn.microsoft.com/zh-tw/library/office/ff745960.aspx
    # Presentation.PrintOptions Property (PowerPoint)
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-printoptions-property-powerpoint
    # PrintOptions.Ranges Property (PowerPoint)
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/printoptions-ranges-property-powerpoint
    # PrintRanges.Add Method (PowerPoint)
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/printranges-add-method-powerpoint
    $printRange = $document.PrintOptions.Ranges.Add(1, $document.Slides.Count)
    
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/ppprintrangetype-enumeration-powerpoint
    # https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.powerpoint.ppprintrangetype(v=office.14).aspx
    # ppPrintAll	Print all slides in the presentation.
    # ppPrintSelection	Print a selection of slides.
    # ppPrintCurrent	Print the current slide from the presentation.
    # ppPrintSlideRange	Print a range of slides.
    # ppPrintNamedSlideShow	Print a named slide show.
    # ppPrintSection	Print the slides of a slideshow section.
    $rangeType = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintAll
    
    # String	The name of the slide show.
    $slideShowName = "Slideshow Name"
    
    # Boolean	Whether the document properties should also be exported. The default is False.
    $includeDocProperties = $false
    
    # Boolean	Whether the IRM settings should also be exported. The default is True.
    $keepIRMSettings = $true
    
    # Boolean	Whether to include document structure tags to improve document accessibility. The default is True.
    $docStructureTags = $true
    
    # Boolean	Whether to include a bitmap of the text. The default is True.
    $bitmapMissingFonts = $true
    
    # Boolean	Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is False.
    $useISO19005_1 = $false
    
    # Boolean	Whether the resulting document should include associated pen marks.
    $includeMarkup = $true
    
    # Variant	A pointer to an Office add-in that implements the IMsoDocExporter COM interface and allows calls to an alternate implementation of code. The default is a null pointer.
    $externalExporter = $null
    
    ##
    ## Publishes as PDF or XPS.
    ##
    ## vba - difference between ExportAsFixedFormat2 and ExportAsFixedFormat? - Stack Overflow
    ## https://stackoverflow.com/questions/37585025/difference-between-exportasfixedformat2-and-exportasfixedformat
    ##
    
    # ExportAsFixedFormat2 can include pen markups
    # Presentation.ExportAsFixedFormat2 Method (PowerPoint)
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-exportasfixedformat2-method-powerpoint
    $document.ExportAsFixedFormat2($exportPath, $fixedFormatType, $intent, $frameSlides, $handoutOrder, $outputType, $printHiddenSlides, $printRange, $rangeType, $slideShowName, $includeDocProperties, $keepIRMSettings, $docStructureTags, $bitmapMissingFonts, $useISO19005_1, $includeMarkup)
    
    # ExportAsFixedFormat cannot include pen markups
    # Presentation.ExportAsFixedFormat Method (PowerPoint)
    # https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-exportasfixedformat-method-powerpoint
    #$document.ExportAsFixedFormat($exportPath, $fixedFormatType, $intent, $frameSlides, $handoutOrder, $outputType, $printHiddenSlides, $printRange, $rangeType, $slideShowName, $includeDocProperties, $keepIRMSettings, $docStructureTags, $bitmapMissingFonts, $useISO19005_1)
    
    # Close PowerPoint file
    $document.Close()
}
# Exit and release the PowerPoint object
$ppt_app.Quit()

# Make sure references to COM objects are released, otherwise powerpoint might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();

