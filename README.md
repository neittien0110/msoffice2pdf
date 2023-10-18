# doc2pdf
Convert all .doc|docx|xlsx files in a directory and its sub to pdf. Store pdf at the same location.

# Usages:

```powershell
doc2pdf.ps1 [[-NameCondition] <string>] [[-Folder] <string>] [<CommonParameters>]
ppt2pdf.ps1 [[-NameCondition] <string>] [[-Folder] <string>] [<CommonParameters>]
xlsx2pdf.ps1 [[-NameCondition] <string>] [[-Folder] <string>] [<CommonParameters>]
```
For examples:

```powershell
doc2pdf.ps1 -NameCondition chapter_  -Folder "C:\Users\Public"
```

![Example](images/0d4482ca2adde917f85be21c25d3933927c93cbcd2b48191820aa36f8b8f790c.png)  

# Note
ps1 is the executable file, so it'll be blocked after downloading from internet. Unblock it to run like the image below.

![unblock the script](images/a6e8977c5bd183f78371e53ed6dea562d53890ab1e6ab1be113b25de49f7ed2a.png)  


# Thanks
Adapted from https://gist.github.com/allenyllee/5d7c4a16ae0e33375e4a6d25acaeeda2
Thank to mp4096, the author of the script ppt2pdf which I customized to this script

and from mp4096
  "Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
   Thanks to MFT, takabanana, ComFreek"

# xlsx2pdf
convert all excel sheets to pdf
NOTE: Change 2 parameters below to force the begin and end page indexes.
    $From = [System.Type]::Missing
    $To = 3
