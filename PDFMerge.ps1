
# old banner warnings about word forced closed

<#
Write-Host "`nPlease note, Word will be forced closed. Please save and/or close documents."
Write-Host "PDF's Will be saved to your Desktop folder"
Write-Host "`nPress any key to continue...`n"
#>

#--->

<#
$bannerMessage = (" " * 4) + @"
Please note, Word will be forced closed. Please save and/or close documents. PDF's Will be saved to your Desktop folder
"@
$bannerMessageWidth = $bannerMessage.Length-1
for ($spc = 0; $spc -lt ($consoleWdth-$bannerMessageWidth); $spc++) {
    $bannerMessage += " "
}

Write-Host $blankLine
Write-Host $blankLine
Write-Host $bannerMessage
Write-Host $blankLine
Write-Host $blankLine
#>

# old banner warning about word forced closed

$consoleWdth = $Host.UI.RawUI.BufferSize.Width-1
$blankLine = " "
for ($spc=0;$spc -lt $consoleWdth;$spc++) {
    $blankLine += " "
}
$pacer = (" " * 4)

$allBannerLines = @(
    $blankLine,
    $blankLine,
    ($pacer + "Please note, Word will be forced closed."),
    ($pacer + "Please save and/or close documents."),
    ($pacer + "PDF's Will be saved to your Desktop folder."),
    $blankLine,
    $blankLine
)

Clear-Host

Write-Host " "
Write-Host " "

$defaultBackground = $Host.UI.RawUI.BackgroundColor
$defaultForeground = $Host.UI.RawUI.ForegroundColor

$Host.UI.RawUI.BackgroundColor = "DarkCyan"
$Host.UI.RawUI.ForegroundColor = "Yellow"

foreach ($bannerLine in $allBannerLines) {
    #$bannerLine = (" " * 4) + $bannerLine
    $bannerLineWidth = $bannerLine.Length-1
    for ($spc = 0; $spc -lt ($consoleWdth-$bannerLineWidth); $spc++) {
        $bannerLine += " "
    }
    Write-Host $bannerLine
}

$Host.UI.RawUI.BackgroundColor = $defaultBackground
$Host.UI.RawUI.ForegroundColor = $defaultForeground

$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null

###--->


function prepFileDetails {
    $userDesktop = [System.Environment]::GetFolderPath("Desktop")
    $response
    $fullFilesPath
    $filesPicked = @()
    $CLIArgs = $args[0]
    if ($CLIArgs) {
        foreach ($CLIArg in $CLIArgs) {
            $filesPicked += (Get-Item -Path $CLIArg).FullName
            $fullFilesPath = Split-Path (Get-Item -Path $CLIArg).FullName

        }
        $response = 'OK'
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        $filePicker = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $true;
            InitialDirectory = $userDesktop
            Title = "Select PDFs and/or RTFs"
            Filter = 'All Files (*.*)|*.*|PDFs (*.pdf)|*.pdf|RTFs (*.rtf)|*.rtf'
        }
        $response = $filePicker.ShowDialog()
        if ($response -eq 'OK') {
            $fullFilesPath = Split-Path $filePicker.FileName
            $filesPicked = $filePicker.FileNames
        } else {
            $PDFMergeError = Get-Content -Path "$PSScriptRoot\PDFMergeError.txt" -Raw

            #Write-Host $PDFMergeError
            Write-Host "Cancelled.`nPress any key to exit.`n"
            $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
            exit
        }
    }

    Write-Host "Cataloging files..."
    $filesMatrix = @()
    foreach ($fullFileName in $filesPicked) {
        $fileName = (Get-Item -Path $fullFileName).BaseName
        $creationTime = ((Get-Item -Path $fullFileName).CreationTime).ToString("yyyy-mm-dd-hh-mm-ss-ms-ms")
        if (($fullFileName.ToLower()).EndsWith('.rtf')) {

            Write-Host "`t$fullFileName"
            $newFileName = "$fullFilesPath\$fileName.pdf"
            $filesMatrix += [PSCustomObject] @{
                fileName = $fileName
                fullFileName = $fullFileName
                creationTime = $creationTime
                newFileName = $newFileName
            }
        } elseif (($fullFileName.ToLower()).EndsWith('.pdf')) {

            Write-Host "`t$fullFileName"
            $filesMatrix += [PSCustomObject] @{
                fileName = $fileName
                fullFileName = $fullFileName
                creationTime = $creationTime
                newFileName = $null
            }
        } else {

            Write-Host "`tInvalid file type (skipped)."
            Continue
        }
    }
    return $filesMatrix
}


###--->


function convertRTFs {
    $filesMatrix = $args[0]
    $tempCheck = 0
    foreach ($allFileDetails in $filesMatrix) {
        if ($allFileDetails -ne $null) {
            $fullFilesPath = Split-Path $allFileDetails.fullFileName
            if (($allFileDetails).newFileName) {
                if (!$tempCheck -and !(Test-Path -Path "$fullFilesPath\RTFs")) {
                    $null = New-Item -Path "$fullFilesPath\RTFs" -ItemType Directory
                    $tempCheck = 1
                }
                $moveFromFileName = ($allFileDetails).fullFileName
                $newPDFileName = ($allFileDetails).newFileName
                $moveToFileName = $fullFilesPath + "\RTFs\" + $allFileDetails.fileName + ".rtf"
                
                Write-Host "Moving:`n`t$moveFromFileName`nto`n`t$moveToFileName"
                Move-Item -Path $moveFromFileName -Destination $moveToFileName
                $MSWord = New-Object -ComObject Word.Application
                $MSWord.Visible = $false
                $newPDFile = $MSWord.Documents.Open($moveToFileName)
                
                Write-Host "Converting RTF:`n`t$moveToFileName`nto PDF:`n`t$newPDFileName"
                $newPDFile.SaveAs([ref] $newPDFileName, [ref] 17)
                $MSWord.Quit()
            }
        }
    }
    foreach ( $ODProc in (Get-CimInstance Win32_Process -Filter "name = 'winword.exe'") ) {
        if (Invoke-CimMethod -InputObject $ODProc -MethodName GetOwner | Where-Object -Property User -eq $env:USERNAME) {
            $null = Stop-Process -Id $ODProc.ProcessId
        }
    }
}


###--->


function mergePDFFiles {
    $filesMatrix = $args[0]
    $userDesktop = [System.Environment]::GetFolderPath("Desktop")
    $currentFileList = @()
    
    <#
    Write-Host "Creating new PDF file name:"
    $outputPDFFile = $userDesktop + "\merged_pdf_" + (get-date -Format "yyyy-MM-dd-hh-mm-ss-ms") + ".pdf"
    Write-Host "`t$outputPDFFile"
    #>
    
    foreach ($allFileDetails in $filesMatrix) {
        if ($allFileDetails -ne $null) {
            $currFileName = ''
            if ($allFileDetails.newFileName) {
                $currFileName = $allFileDetails.newFileName
            } else {
                $currFileName = $allFileDetails.fullFileName
            }
            $fileSize = (Get-Item -Path $currFileName).Length
            $fileSizeTrack += $fileSize
            $trackRounded = [math]::ceiling($fileSizeTrack/1MB)
            $currentFileList += $currFileName
            if ($trackRounded -gt 15) {
                
                Write-Host "Creating new PDF file name:"
                $outputPDFFile = $userDesktop + "\merged_pdf_" + (get-date -Format "yyyy-MM-dd-hh-mm-ss-ms") + ".pdf"
                Write-Host "`t$outputPDFFile"
                
                Write-Host "Merging PDF's"
                Merge-PDF -InputFile ($currentFileList) -OutputFile $outputPDFFile
                $currentFileList = @()
                $fileSizeTrack = 0
            }
        }
    }
    
    Write-Host "Creating new PDF file name:"
    $outputPDFFile = $userDesktop + "\merged_pdf_" + (get-date -Format "yyyy-MM-dd-hh-mm-ss-ms") + ".pdf"
    Write-Host "`t$outputPDFFile"
    
    Write-Host "Merging PDF's..."
    Merge-PDF -InputFile ($currentFileList) -OutputFile $outputPDFFile
}


###--->


$filesMatrix = @()
if ($args) {
    $filesMatrix = prepFileDetails($args)
} else {
    $filesMatrix = prepFileDetails
}
convertRTFs($filesMatrix)
mergePDFFiles($filesMatrix)

Write-Host "`nDone! Press any key to exit.`n"
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
