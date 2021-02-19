Attribute VB_Name = "modTest"
Option Explicit

Sub MakeSampleSheet()

    Dim sh As Worksheet
    For Each sh In ActiveWindow.SelectedSheets
        sh.Select
        Exit For
    Next
       
    If WorksheetFunction.CountA(Range(Cells(1, 1), Cells(25, 7))) > 0 Then
        MsgBox "The sample sheet area Range(Cells(1, 1), Cells(25, 7)) is not empty. Select empty sheet."
        Exit Sub
    End If


    Cells(1, 1) = "FLG"
    Cells(1, 2) = "Picture"
    Cells(1, 3) = "Text"
    Cells(4, 1) = "Text from below"
    Cells(5, 3) = ":sectnums:"
    Cells(8, 1) = "'=="
    Cells(8, 3) = "Excel to Asciidoctor"
    Cells(10, 1) = "'==="
    Cells(10, 3) = "Ordinary sentence"
    Cells(12, 3) = "This is a normal sentence."
    Cells(14, 1) = "'==="
    Cells(14, 3) = "Table"
    Cells(16, 1) = "table"
    Cells(16, 3) = "Header1"
    Cells(16, 4) = "Header2"
    Cells(16, 5) = "Header3"
    Cells(16, 6) = "Header4"
    Cells(16, 7) = "Header5"
    Cells(17, 3) = "Data11"
    Cells(17, 4) = "Data12"
    Cells(17, 5) = "Data13"
    Cells(17, 6) = "Data14"
    Cells(17, 7) = "Data15"
    Cells(18, 3) = "Data21"
    Cells(18, 4) = "Data22"
    Cells(18, 5) = "Data23"
    Cells(18, 6) = "Data24"
    Cells(18, 7) = "Data25"
    Cells(19, 3) = "Data22"
    Cells(19, 4) = "Data23"
    Cells(19, 5) = "Data24"
    Cells(19, 6) = "Data25"
    Cells(19, 7) = "Data26"
    Cells(20, 1) = "table"
    Cells(20, 3) = "Data23"
    Cells(20, 4) = "Data24"
    Cells(20, 5) = "Data25"
    Cells(20, 6) = "Data26"
    Cells(20, 7) = "Data27"
    Cells(23, 1) = "'==="
    Cells(23, 3) = "Picture"
    Cells(25, 2) = "$B$31.png"
    
    Call ActiveSheet.Shapes.AddShape(msoShapeRectangle, 120.6, 460.8, 227.4, 103.2)
End Sub
