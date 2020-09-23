Attribute VB_Name = "modMain"
' API declarations
Declare Function AddFontResource Lib "GDI32" Alias "AddFontResourceA" (ByVal FontFileName As String) As Long
Declare Function RemoveFontResource Lib "GDI32" Alias "RemoveFontResourceA" (ByVal FontFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Const MAX_PATH = 260

Public Function GetWinPath()
Dim strFolder As String
Dim lngResult As Long

strFolder = String(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
Else
    GetWinPath = ""
End If
End Function



