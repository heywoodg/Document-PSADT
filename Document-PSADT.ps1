Function Get-PSADTCode ($sourcefile, $searchstart, $searchend) {
    $content = New-Object System.Collections.ArrayList
    ForEach ($line in $sourcefile) {
        If ($line -like "$searchstart") {
            $writeit = $true
        }
        if ($writeit) {
            if ($line -like "$searchend") {
                $writeit = $false
            }
            Else {
                $content.Add($line) >$null
            }
        }
    }
    Return $content
}
Function Write-PSADTCode ($content) {
    $SectionLines = ($content.length)
    $lineNum = 0
    Do {
        $selection.Style = "No Spacing"
        If ($content[$LineNum] -like "*#*") {
            $selection.font.color = 13056                   
        }
        Else {
            $selection.font.color = 8388608
        }
        $selection.font.name = "terminal"
        $selection.font.size = 10
        $selection.typeText(  ($content[$LineNum]).trimstart()   )
        $selection.TypeParagraph()
        $LineNum++ 
    } Until ($LineNum -eq $SectionLines)
    $selection.TypeParagraph()
    $selection.TypeParagraph()
}

#set variables
$date = (get-date).ToShortDateString()
$filePath = "C:\temp\Deploy-Application.ps1"
$Wordfile = "c:\temp\document.doc"
$PS1file = Get-Content -Path $filepath

#Setup the Word document
[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$word = New-Object -ComObject word.application
$word.visible = $true
$doc = $word.documents.add()
$banner = ($pwd.path + "\AppDeployToolkit\AppDeployToolkitBanner.png")

# Variables Section  
$searchstart = '*## Variables: Application*'
$searchend = '*##*===============================================*'
$vars = New-Object System.Collections.ArrayList
$vars = get-PSADTCode $PS1file $searchstart $searchend
$vendor = ([string]($vars | select-string "appvendor")).split("'")[-2]
$appname = ([string]($vars | select-string "appname")).split("'")[-2]
$version = ([string]($vars | select-string "appversion")).split("'")[-2]

#Set the format and heading information
$selection = $word.selection
$selection.WholeStory
$Selection.InlineShapes.AddPicture("$banner")
$selection.TypeParagraph()
$selection.Style = "Title"
$selection.typeText("PSADT: $vendor $appname $version Deployment")
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.font.size = 12
$selection.typeText("Date: $date")
$selection.TypeParagraph()
$selection.insertnewpage()

#Set up Table of Contents
$selection.font.size = 14
$selection.font.bold=$True
$selection.typeText("Table of Contents")
$selection.TypeParagraph()
$range = $selection.range
$toc = $doc.TablesOfContents.Add($Range)
$selection.TypeParagraph()
$selection.insertnewpage()

# Write overview information
$text = @"
This document provides details of the Powershell Application Deployment Toolkit (PSADT) `
configuration used for the deployment of $vendor $appname $version. Details of the variables used `
as well as the installation and uninstallation configuration are included.
"@
$selection.Style = "Heading 1"
$selection.typeText("Overview")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Write variable information to document
$selection.Style = "Heading 1"
$selection.typeText("Variables")
$selection.TypeParagraph()
$selection.Style = "Normal"
$selection.typeText("Below are the variables defined in the deploy-application.ps1 file. ")
$selection.TypeParagraph()
Write-PSADTCode $vars
$selection.TypeParagraph()
$selection.TypeParagraph()

# Pre-Installation Section
$searchstart = '*Pre-Installation*'
$searchend = '*##*===============================================*'
$preinst = New-Object System.Collections.ArrayList
$preinst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Pre-Installation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Pre-Installation section of the script will store any commands that have been added, and `
are intended to run before the main installation. Typically these commands could be required `
to set up the environment or close applications before the main installation can run. 

Pre-requisites (such as Abode Flash, MS Visual C++ Runtime, etc., should not be done here, but `
should be done from within SCCM.
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $preinst
$selection.TypeParagraph()
$selection.TypeParagraph()

# Installation Section
$searchstart = '*Perform Installation tasks here*'
$searchend = '*##*===============================================*'
$inst = New-Object System.Collections.ArrayList
$inst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Installation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Installation section of the script will store any commands that have been added to `
perform the actual installation of the software package. 
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $inst
$selection.TypeParagraph()
$selection.TypeParagraph()

# Post-installation Section
$searchstart = '*Perform Post-Installation tasks here*'
$searchend = '*##*===============================================*'
$postinst = New-Object System.Collections.ArrayList
$postinst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Post-Installation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Post-installation section of the script will store any commands that have been added to `
be run after the installation of the software package. 
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $postinst

# Pre-uninstallation Section
$searchstart = "*$installPhase = 'Pre-Uninstallation'*"
$searchend = '*##*===============================================*'
$preuninst = New-Object System.Collections.ArrayList
$preuninst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Pre-Unstallation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Pre-Uninstallation section of the script will store any commands that have been added to `
be run bfore an installation attempt is made of the software package. This normally includes a `
command to to close the actual application if it is running.
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $preuninst

# Uninstallation Section
$searchstart = "*$installPhase = 'Uninstallation'*"
$searchend = '*##*===============================================*'
$uninst = New-Object System.Collections.ArrayList
$uninst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Unstallation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Uninstallation section of the script will store any commands that need to be rune run to do `
the installation of the software package.
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $uninst

# Post-Uninstallation Section
$searchstart = "*$installPhase = 'Post-Uninstallation'*"
$searchend = '*##*===============================================*'
$Postuninst = New-Object System.Collections.ArrayList
$Postuninst = get-PSADTCode $PS1file $searchstart $searchend
$selection.Style = "Heading 1"
$selection.typeText("Post-Unstallation")
$selection.TypeParagraph()
$selection.Style = "No Spacing"
$text = @"
The Post-Uninstallation section of the script will store any commands that need to be rune run after the`
uninstallation of the software package. This could be to do some kind of clean-up. If a replacement product `
is to be installed, that should be done using SCCM superscedence.
"@
$selection.typeText($text)
$selection.TypeParagraph()
$selection.TypeParagraph()
Write-PSADTCode $Postuninst

#Update TOC
$toc.Update()

#Add End of Document page
$selection.insertnewpage()
$i=0
Do {
    $selection.TypeParagraph()
    $i++

} While ($i -lt 15)
$selection.Style = "No Spacing"
$Selection.ParagraphFormat.Alignment = 1
$selection.font.size = 20
#$selection.font.bold=$True
$selection.typeText("END OF DOCUMENT")
$selection.TypeParagraph()
$selection.typeText("INTENTIONALLY BLANK")
$selection.TypeParagraph()

# Save the document
$doc.saveas([ref] $Wordfile, [ref]$saveFormat::wdFormatDocument)


