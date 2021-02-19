Attribute VB_Name = "modCommon"
Option Explicit

Public Function GetFileExtension(ByVal FileName As String, Optional ByRef errMsg As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo EH
    GetFileExtension = fso.GetExtensionName(FileName)
    GoTo NE
EH:
    GetFileExtension = ""
    errMsg = "GetFileExtension error."
NE:
    Set fso = Nothing
End Function
