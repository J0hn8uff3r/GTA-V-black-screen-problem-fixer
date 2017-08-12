Option Explicit

Const ForReading = 1
Const ForWriting = 2
Dim blnRes: blnRes = 0
Dim wrongStr, goodStr, path
wrongStr = "<DX_Version value=""2"" />"
goodStr = "<DX_Version value=""0"" />"
path = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\Rockstar Games\GTA V"

On Error Resume Next
Call update(path & "\settings.xml", wrongStr, goodStr)
blnRes = blnRes Or (Err.Number = 0): Err.Clear
On Error GoTo 0

If blnRes Then MsgBox("GTA V black screen problem successfully fixed!") Else MsgBox("Something went wrong!") End If

Sub update(strFile, wrongStr, goodStr)
Dim objFSO, objFile, strText, goodStrText
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFile, ForReading)
strText = objFile.ReadAll
objFile.Close
goodStrText = Replace(strText, wrongStr, goodStr)
Set objFile = objFSO.OpenTextFile(strFile, ForWriting)
objFile.WriteLine goodStrText
objFile.Close
End Sub
