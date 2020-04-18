Attribute VB_Name = "onebitRL"
'@Folder("Main")
Option Explicit

Private Sub MakeSquareCells()

    Dim cursheet As Worksheet

    Set cursheet = ActiveSheet

    With cursheet
        .Columns.ColumnWidth = 1
        .Rows.EntireRow.RowHeight = .Cells(1).Width
    End With

End Sub

Private Sub ArrayToWorksheet(ByVal varArray As Variant, ByRef ws As Worksheet)
    Dim range As range
    
    Set range = ws.range("A1")
    
    Debug.Assert range.Address = "$A$1"
    
    Set range = range.Resize(UBound(varArray, 1), UBound(varArray, 2))
    
    range.Value2 = varArray

End Sub

Private Sub testhex()
    Dim value As Long
    Dim str As String
    
    str = "FF"
    value = "&H" & str
    
    Debug.Print value

End Sub
