Attribute VB_Name = "modAsciidoc"
Option Explicit

Private Const DATA_START_ROW As Long = 5
Private Const FLAG_COL As Long = 1
Private Const PIC_COL As Long = 2
Private Const VALUE_COL As Long = 3
Private Const TABLE_ROW_MAX As Long = 1000
Private Const TABLE_COL_MAX As Long = 20


Public Sub MakeDocumentAndPic()
    Dim myLogs As clsLogs
    Set myLogs = New clsLogs
    
    Call MakeDocument
    ActiveSheet.Shapes.SelectAll
    Call SavePictures(myLogs)
    myLogs.OutputErrs "", True
End Sub

Public Sub MakeDocument()
    Cells(1, 1).Select
    Dim rngData As Range
    Dim rowEnd As Long
    Set rngData = Application.Intersect(Range(Cells(1, FLAG_COL), Cells(Rows.Count, VALUE_COL)), ActiveSheet.UsedRange)
    rowEnd = rngData.Row + rngData.Rows.Count
        
    Dim i As Long
    Dim asciidoc As clsString
    Set asciidoc = New clsString
    
    For i = DATA_START_ROW To rowEnd
        asciidoc.Add MakeBody(i) & vbCrLf
    Next i
    
    asciidoc.SaveToFileUTF8 ActiveWorkbook.Path & "\" & ActiveWorkbook.ActiveSheet.Name & ".adoc"
End Sub

Private Function MakeBody(ByRef iRow As Long) As String
    Dim flg As String
    Dim picValue  As String
    Dim textValue As String

    Dim i As Long
    Dim j As Long
    flg = Trim((Cells(iRow, FLAG_COL)))
    picValue = Trim((Cells(iRow, PIC_COL)))
    textValue = Trim(Cells(iRow, VALUE_COL))
    
    Dim tempBody As String
    
    If picValue <> "" Then
        tempBody = "image::" & picValue & "[]"
        
    ElseIf LCase(flg) = "table" Then
        Dim tableValues As clsString
        Dim values As clsString
        Set tableValues = New clsString
        Set values = New clsString
        
        Dim tableColumns() As Long
        Dim tableMaxCol As Long
        tableMaxCol = GetTableColumns(iRow, VALUE_COL, 25, tableColumns)
        If tableMaxCol > TABLE_COL_MAX Then
            MsgBox "col is over " & TABLE_COL_MAX
            End
        End If
            
        '' header row
        tableValues.Add "|==="
        For j = VALUE_COL To VALUE_COL + tableMaxCol - 1
            values.Add "a|" & Cells(iRow, tableColumns(j - VALUE_COL))
        Next j
        tableValues.Add values.Joins(" ")
           
        'data
        For i = iRow + 1 To iRow + TABLE_ROW_MAX
            For j = VALUE_COL To VALUE_COL + tableMaxCol - 1
                tableValues.Add "a|" & Cells(i, tableColumns(j - VALUE_COL))
            Next j
            tableValues.Add ""
            
            If LCase(Cells(i, FLAG_COL)) = "table" Then
                tableValues.Add "|==="
                iRow = i + 1
                Exit For
            End If
        Next i
        
        ' rows check
        If i >= iRow + TABLE_ROW_MAX Then
            MsgBox "table row is over " & TABLE_ROW_MAX
            End
        End If
        
        tempBody = tableValues.Joins(vbCrLf) & vbCrLf
    Else
        If flg <> "" Then
            tempBody = flg & " " & textValue
        Else
            tempBody = textValue
        End If
    End If
    
    MakeBody = tempBody
End Function

Public Sub SavePictures(ByRef logs As clsLogs)
    On Error GoTo SavePictures_Error
        
    Dim iCount As Long
    On Error Resume Next
    iCount = Selection.ShapeRange.Count
    If Err.Number <> 0 Then
        iCount = 0
    End If
    
    On Error GoTo SavePictures_Error
    If iCount = 0 Then
        If MsgBox("Save all images?", vbOKCancel) = vbOK Then
            ActiveSheet.Shapes.SelectAll
        Else
            GoTo normalEx
        End If
    End If
    
    Dim pics As ShapeRange
    Set pics = Selection.ShapeRange
    
    pics.Item(1).TopLeftCell.Select
    
    Dim FileName As String
    Dim savePath As String
    Dim ext As String
    savePath = ActiveWorkbook.Path
    Dim sp As Shape
    For Each sp In pics
        If sp.Type <> msoComment Then
            FileName = sp.TopLeftCell.Offset(-1, -1)
            ext = GetFileExtension(FileName)
           
            ' make file name
            Dim retCode As Long
            If FileName = "" Then
                FileName = replace(sp.TopLeftCell.Offset(-1, -1).Address, ":", "") & ".png"
                sp.TopLeftCell.Offset(-1, -1) = FileName
            ElseIf ext = "" Then
                FileName = FileName & ".png"
                sp.TopLeftCell.Offset(-1, -1) = FileName
            End If
            
            ' save picture
            sp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
            retCode = SaveClipBoard(savePath & "\" & FileName)
            If retCode <> 0 Then
                logs.AddErr ActiveSheet.Name, sp.TopLeftCell.Address, "Picture Save error."
            End If
        End If
    Next
normalEx:
    On Error GoTo 0
    Exit Sub
SavePictures_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SavePictures, line " & Erl & "."
End Sub


Private Function GetTableColumns(ByVal r As Long, ByVal c As Long, ByVal maxCols As Long, columns() As Long) As Long
    Dim i As Long
    ReDim columns(0 To maxCols)
    Dim rng As Range
    Set rng = Cells(r, c)
    For i = 0 To maxCols
        If rng.text = "" Then
            GetTableColumns = i
            Exit Function
        Else
            columns(i) = rng.MergeArea.Column
        End If
        Set rng = rng.Offset(0, 1)
    Next i
End Function

