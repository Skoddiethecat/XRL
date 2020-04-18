VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameInitalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Classes")
Option Explicit

Private AltairLib As AltairLib
Private GameData As GameData
Private Renderer As Renderer

Private Sub Class_Initialize()
    Set AltairLib = alFactory.AltairLibLoad
    Set GameData = Factories.GameDataCreate
    Set Renderer = Factories.RendererCreate
    
End Sub

Public Sub Start(ByRef GameSheet As Worksheet)
    AltairLib.alOptimizer.Enable DisplayPageBreaks:=False
    
    LoadTiles
    
    CreateScreen GameSheet:=GameSheet
    
    Renderer.SetGameSheet ws:=GameSheet
    
    Renderer.Draw
    
    AltairLib.alOptimizer.Disable DisplayPageBreaks:=False
    
End Sub

Public Sub Halt(ByRef GameSheet As Worksheet)
    AltairLib.alOptimizer.Enable DisplayPageBreaks:=False
    
    Application.DisplayAlerts = False
    
    GameSheet.Delete
    
    Application.DisplayAlerts = True
    
    AltairLib.alOptimizer.Disable DisplayPageBreaks:=False

End Sub

Private Sub LoadTiles()
    Dim TileSheet As Worksheet
    Dim TileArray() As Variant
    Dim i As Long
    
    Set TileSheet = ThisWorkbook.Worksheets("Tiles")
    
    TileArray = TileSheet.UsedRange.Value2
    
    For i = LBound(TileArray, 1) To UBound(TileArray, 1)
        GameData.Tiles.AddTile value:=TileArray(i, 1), Key:=CStr(TileArray(i, 2))
    
    Next i
    

End Sub

Private Sub CreateScreen(ByRef GameSheet As Worksheet)
    Dim Screen As range
    
    Set GameSheet = ThisWorkbook.Worksheets.Add
    
    With GameSheet
        .Name = "GameSheet"
        .DisplayPageBreaks = False
        .Columns.ColumnWidth = 1
        .Rows.EntireRow.RowHeight = .Cells(1).Width
        .Cells.Interior.Color = TrueColor
    
    End With
    
    Set Screen = GameSheet.Cells(1, 1)
    
    Set Screen = Screen.Resize(ScreenResolutionRows, ScreenResolutionColumns)
    
    'Screen.Interior.Color = FalseColor
    
    ActiveWindow.Zoom = 20

End Sub

Sub TileDebug(Key As String)

    Dim i As Long
    Dim j As Long
    Dim str As String
        
    For i = 0 To 7
        str = ""
        
        For j = 0 To 7
            If GameData.Tiles.GetTile(Key)(i, j) Then
                str = str & "#"
            
            Else
                str = str & " "
            
            End If
            
            
            
        Next j
        
        Debug.Print str
    
    Next i
    
End Sub
