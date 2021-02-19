VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCollection As Collection

Private Sub Class_Initialize()
    Set mCollection = New Collection
End Sub

Public Sub Add(ByRef rData As Variant)
    If StrPtr(rData) = 0 Then Exit Sub
    mCollection.Add CStr(rData)
End Sub

Public Function Count() As Long
    Count = mCollection.Count
End Function

Public Function Joins(Optional ByVal vSep As String = vbCrLf) As String
    If mCollection.Count = 0 Then
        Joins = vbNullString
        Exit Function
    End If
    Joins = Join(Me.MakeArray, vSep)
End Function
Public Function MakeArray() As String()
    Dim myStrings() As String
    Dim i As Long

    If mCollection.Count > 0 Then
        ReDim myStrings(0 To mCollection.Count - 1)
        For i = 0 To UBound(myStrings)
            myStrings(i) = mCollection.Item(i + 1)
        Next i
    Else
        ReDim myStrings(0 To 0)
    End If
    MakeArray = myStrings
End Function

Public Function SaveToFileUTF8(fn As String) As Boolean
    Dim writeStream As Object
    On Error GoTo saveutf8_Error

    If mCollection.Count = 0 Then
        Exit Function
    End If

    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open
    
    Dim i As Long
    For i = 1 To mCollection.Count
        writeStream.WriteText mCollection.Item(i) ' 1: add new line
    Next i

    writeStream.Position = 0
    writeStream.Type = 1 'adTypeBinary:1    default: 2
    writeStream.Position = 3

    Dim byteData() As Byte
    byteData = writeStream.Read
    writeStream.Close

    writeStream.Open
    writeStream.Write byteData

    writeStream.SaveToFile fn, 2 'adSaveCreateOverWrite:2

    On Error GoTo 0
    GoTo saveutf8_Normal_exit

saveutf8_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveutf8 of Class Module clsStrings"

saveutf8_Normal_exit:
    writeStream.Close
    Set writeStream = Nothing
End Function




