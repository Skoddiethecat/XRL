VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Classes.GameData")
Option Explicit

Private Tiles As Collection
Private overwrite As Boolean

Private Sub Class_Initialize()
    Set Tiles = New Collection
    overwrite = False

End Sub

Public Sub AddTile(ByVal value As String, ByVal Key As String)
    Dim Tile(0 To 7, 0 To 7) As Boolean
    Dim RowData As String
    Dim MsgBoxResults As Long
    Dim i As Long
    Dim j As Long
        
    If TileErrorChecking(Key) Then Exit Sub
    
    For i = 0 To 7
        RowData = Mid(value, (i * 2) + 1, 2)
        RowData = Application.WorksheetFunction.Hex2Bin(RowData, 8)
        
        For j = 0 To 7
            If Mid(RowData, j + 1, 1) = 1 Then
                Tile(i, j) = True
            
            Else
                Tile(i, j) = False
                
            End If
        
        Next j
    
    Next i

    Tiles.Add Tile, Key

End Sub

Public Function GetTile(ByVal Key As String, ByVal row As Long, column As Long) As Boolean
    GetTile = Tiles(Key)(row, column)

End Function

Private Function TileErrorChecking(ByVal Key As String) As Boolean
    Dim MsgBoxResults As Long
    TileErrorChecking = True

    If KeyExists(Tiles, Key) Then
        If Not overwrite Then
        
            MsgBoxResults = MsgBox(Prompt:="Duplicate key found, overwrite all duplicate tiles?", _
                                   Buttons:=vbYesNoCancel, Title:="Duplicate Tile Key!")
            
            Select Case MsgBoxResults
                Case vbYes
                    overwrite = True 'overwrite this tile and all others
                
                Case vbNo
                    overwrite = False 'overwrite this time, but not all other
                
                Case vbCancel
                    Exit Function 'Cancel and check again
                
            End Select
            
        End If
        
        Tiles.Remove Key
    
    End If
    
    TileErrorChecking = False

End Function

Private Function KeyExists(ByRef col As Collection, ByVal Key As String) As Boolean
    Dim Element As Variant
    
    On Error Resume Next
    
    Element = col(Key)
    KeyExists = (Err.number = 0)
    
    Err.Clear
    
    On Error GoTo 0
    
End Function
