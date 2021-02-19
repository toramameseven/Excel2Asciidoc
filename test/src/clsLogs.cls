VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private logs As Collection
Private errs As Collection

Private Sub Class_Initialize()
    Set logs = New Collection
    Set errs = New Collection
End Sub

Public Sub AddLog(ByVal sheetName As String, ByVal address1 As String, ByVal msg1 As String)
    logs.Add msg1 & "|" & address1 & "|" & sheetName
End Sub

Public Sub AddErr(ByVal sheetName As String, ByVal address1 As String, ByVal msg1 As String)
    errs.Add msg1 & "|" & address1 & "|" & sheetName
End Sub

Public Sub OutputLogs(Optional ByVal txtFormat = "|1", Optional ByVal isHyperlink As Boolean = False)
    If txtFormat = "" Then
        txtFormat = "|1"
    End If
    OutPutCommon 0, txtFormat, isHyperlink
End Sub

Public Sub OutputErrs(Optional ByVal txtFormat = "|3-|2 : |1", Optional ByVal isHyperlink As Boolean = True)
    If errs.Count = 0 Then Exit Sub
    If txtFormat = "" Then
        txtFormat = "|3-|2 : |1"
    End If
    OutPutCommon 1, txtFormat, isHyperlink
End Sub


'############################################################################################

Private Sub OutPutCommon(ByVal logType As Long, ByVal txtFormat, ByVal isHyperlink As Boolean)
    Dim sheetName As String
    Dim logClass As Collection

    If logType = 0 Then
        sheetName = "logs"
        Set logClass = logs
    ElseIf logType = 1 Then
        sheetName = "errs"
        Set logClass = errs
    Else
        Exit Sub
    End If

    If isExistSheet(sheetName) Then
        '
    Else
        ActiveWorkbook.Sheets.Add.Name = sheetName
    End If

    '' output errors and logs
    Dim ii As Long

    '' output errs to worksheet
    Worksheets(sheetName).Cells.Clear
    Dim msgList() As String
    Dim dispText As String
    For ii = 1 To logClass.Count
        msgList = Split(logClass.Item(ii), "|")

        dispText = replace(txtFormat, "|1", msgList(0))
        dispText = replace(dispText, "|2", msgList(1))
        dispText = replace(dispText, "|3", msgList(2))

        If isHyperlink Then
            ActiveSheet.Hyperlinks.Add anchor:=Worksheets(sheetName).Cells(ii, 1), Address:="", SubAddress:="'" & msgList(2) & "'!" & msgList(1), TextToDisplay:=dispText
        Else
            Worksheets(sheetName).Cells(ii, 1) = dispText
        End If
    Next

    If errs.Count > 0 Then
        Worksheets(sheetName).Select
        Cells(1, 1).Select
    End If
End Sub

Private Function isExistSheet(ByVal sheetName) As Boolean
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        If UCase(st.Name) = UCase(sheetName) Then
            isExistSheet = True
            Exit Function
        End If
    Next
    isExistSheet = False
End Function





