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
#$xlsx_app = New-Object -ComObject PowerPoint.Application
#
# or
#
##
## Using Interop Assemblies | Richard Siddaway's Blog
## https://richardspowershellblog.wordpress.com/2007/05/29/using-interop-assemblies/
##
# Load Powerpoint Interop Assembly
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Excel") > $null
[Reflection.Assembly]::LoadWithPartialname("Office") > $null # need this or powerpoint might not close
$xlsx_app = New-Object -ComObject excel.application

# Show the Microsoft Excel app, or not
$xlsx_app.visible = $true

# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.xls? | ForEach-Object {
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
    $workbook = $xlsx_app.workbooks.Open($_.FullName)
        
    ##
    ## .net - powershell - extract file name and extension - Stack Overflow
    ## https://stackoverflow.com/questions/9788492/powershell-extract-file-name-and-extension
    ##
    ## If the file is coming off the disk and as others have stated, use the BaseName and Extension properties
    ##
    # Create a name for the PDF workbook; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    #$pdf_filename = "$($curr_path)\$($_.BaseName).pdf"
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    
    ##
    ## Presentation.SaveAs Method (PowerPoint)
    ## https://msdn.microsoft.com/en-us/vba/powerpoint-vba/articles/presentation-saveas-method-powerpoint
    ##
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    #$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    #$workbook.SaveAs($pdf_filename, $opt)
    
    
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
    $fixedFormatType = [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF


    # Can be set to either xlQualityStandard or xlQualityMinimum.
    $Quality = [Microsoft.Office.Interop.Excel.XlFixedFormatQuality]::xlQualityStandard
    
    # Set to True to indicate that document properties should be included or set to False to indicate that they are omitted.
    $IncludeDocProperties = $true
    
    # If set to True, ignores any print areas set when publishing. If set to False, will use the print areas set when publishing.
    $IgnorePrintAreas = $false
    	
    # The number of the page at which to start publishing. 
    # If this argument is omitted ([System.Type]::Missing), publishing starts at the beginning.
    $From = [System.Type]::Missing
    
    # The number of the last page to publish. E.g 1,2,3,4,5...
    # If this argument is omitted ([System.Type]::Missing), publishing ends with the last page
    $To = [System.Type]::Missing

    # If set to True displays file in viewer after it is published. If set to False the file is published but not displayed.
    $OpenAfterPublish=$true

    ##
    ## Publishes as PDF or XPS.
    # 
    # https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff198122(v=office.14)
    $workbook.ExportAsFixedFormat($fixedFormatType, $exportPath, $Quality, $IncludeDocProperties, $IgnorePrintAreas, $From, $To, $OpenAfterPublish);
    
    # Close Excel file
    $workbook.Close()

  
}
# Exit and release the PowerPoint object
$xlsx_app.Quit()

# Make sure references to COM objects are released, otherwise powerpoint might not close
# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();



