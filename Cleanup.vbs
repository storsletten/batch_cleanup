Option Explicit

Dim fso, wshShell, scriptPath, scriptName, scriptFileName, pathsFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

scriptFileName = WScript.ScriptName
scriptName = Left(scriptFileName, InStrRev(scriptFileName, ".") - 1)

scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

pathsFile = fso.BuildPath(scriptPath, scriptName & " paths.txt")

If Not fso.FileExists(pathsFile) Then
  Dim ts
  Set ts = fso.CreateTextFile(pathsFile, True)
  ts.WriteLine "# Add paths to folders or files that should be cleaned up."
  ts.WriteLine "# Put each path on a separate line in this file."
  ts.WriteLine "# The paths can be either absolute or relative to the folder where this file is located."
  ts.WriteLine ""
  ts.Close
End If

Dim paths(), line, trimmedLine, count
count = 0
ReDim paths(0)
Dim tsRead
Set tsRead = fso.OpenTextFile(pathsFile, 1)

Do Until tsRead.AtEndOfStream
  line = tsRead.ReadLine
  trimmedLine = Trim(line)
  If Len(trimmedLine) > 0 Then
    If Not (Left(trimmedLine, 1) = "#" Or Left(trimmedLine, 1) = ";" Or Left(trimmedLine, 2) = "//") Then
      If (Left(trimmedLine, 1) = """" And Right(trimmedLine, 1) = """") Then
        If Len(trimmedLine) > 2 Then
          trimmedLine = Mid(trimmedLine, 2, Len(trimmedLine) - 2)
        Else
          trimmedLine = ""
        End If
      End If
      If Len(trimmedLine) > 0 Then
        If count = 0 Then
          paths(0) = trimmedLine
        Else
          ReDim Preserve paths(count)
          paths(count) = trimmedLine
        End If
        count = count + 1
      End If
    End If
  End If
Loop
tsRead.Close

If count = 0 Then
  On Error Resume Next
  Dim ret
  ret = wshShell.Run("notepad.exe """ & pathsFile & """", 1, False)
  If Err.Number <> 0 Then
    MsgBox "The paths file is empty."
  End If
  WScript.Quit
End If

Dim i, currentPath
For i = 0 To UBound(paths)
  currentPath = paths(i)
  If fso.FolderExists(currentPath) Then
    Dim folder, file, subfolder
    Set folder = fso.GetFolder(currentPath)
    For Each file In folder.Files
      On Error Resume Next
      fso.DeleteFile file.Path, True
      On Error GoTo 0
    Next
    For Each subfolder In folder.SubFolders
      On Error Resume Next
      fso.DeleteFolder subfolder.Path, True
      On Error GoTo 0
    Next
  ElseIf fso.FileExists(currentPath) Then
    On Error Resume Next
    fso.DeleteFile currentPath, True
    On Error GoTo 0
  End If
Next

MsgBox "Cleanup complete."
