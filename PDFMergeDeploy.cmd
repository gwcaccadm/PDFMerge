:checkFolder

if not exist "C:\ProgramData\AutoPilotConfig\PDFMerge" (
    mkdir "C:\ProgramData\AutoPilotConfig\PDFMerge"
    set /a CheckFolderResult = 1
)

if %*CheckFolderResult*% EQU 0 goto checkFolder

rem copy "PDFMerge.ico" "C:\ProgramData\AutoPilotConfig\Icons" /y
copy "PDFMerg*.*" "C:\ProgramData\AutoPilotConfig\PDFMerge" /y

:checkFile

if exist "C:\ProgramData\AutoPilotConfig\PDFMerge\PDFMergeShortcut.ps1" (
    Powershell.exe -Executionpolicy bypass -File "C:\ProgramData\AutoPilotConfig\PDFMerge\PDFMergeShortcut.ps1"
    set /a CheckFileResult = 1

) Else (
    copy "PDFMergeShortcut.ps1" "C:\ProgramData\AutoPilotConfig\PDFMerge" /y
    set /a CheckFileResult = 0
)

if %*CheckFileResult*% EQU 0 goto checkFile

if (%CheckFolderResult% + %CheckFileResult%) EQU 2 (
    rem return success code
    echo 1707
)
