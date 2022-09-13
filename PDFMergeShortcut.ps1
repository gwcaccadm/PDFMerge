
$UserDesktop = "C:\Users\Public\Desktop"
$PDFMergeLK = $UserDesktop + "\PDFMerge.lnk"
$AutoPilotDir = "C:\ProgramData\AutoPilotConfig"

if (!(Test-Path $PDFMergeLK)) {
    $PDFMergeSC = (New-Object -ComObject WScript.Shell).CreateShortcut($PDFMergeLK)
    $PDFMergeSC.TargetPath = "$AutoPilotDir\PDFMerge\PDFMerge.cmd"
    $PDFMergeSC.IconLocation = "$AutoPilotDir\PDFMerge\PDFMerge.ico"
    $PDFMergeSC.Description = "Wathaurong RTF/PDF Merging App"
    $PDFMergeSC.WindowStyle = 0
    #$PDFMergeSC.WindowStyle = 7
    #$PDFMergeSC.WorkingDirectory = $UserDesktop
    $PDFMergeSC.WorkingDirectory = [System.Environment]::GetFolderPath("Desktop")
    $PDFMergeSC.Save()
}
