VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Renderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Classes")
Option Explicit

Private GameData As GameData
Private Tiles As Tiles
Private PlayField As PlayField
Private Screen(0 To (ScreenResolutionRows - 1), 0 To (ScreenResolutionColumns - 1)) As Boolean
Private GameSheet As Worksheet

Private Sub Class_Initialize()
    Set GameData = Factories.GameDataCreate
    Set Tiles = GameData.Tiles
    Set PlayField = GameData.PlayField
    
    InitalizeScreen arr:=Screen

End Sub

Public Sub SetGameSheet(ws As Worksheet)
    Set GameSheet = ws

End Sub

Public Sub Draw()
    PlayFieldtoScreen
    ScreentoSheet ws:=GameSheet

End Sub

Private Sub PlayFieldtoScreen()
    Dim i As Long
    Dim j As Long
    Dim TileCol As Long
    Dim TileRow As Long
    Dim TileKey As String
    
    For i = LBound(Screen, 1) To UBound(Screen, 1)
        TileRow = Int(i / 8)
        
        For j = LBound(Screen, 2) To UBound(Screen, 2)
            TileCol = Int(j / 8)
            
            TileKey = PlayField.GetTile(TileRow, TileCol)
            
            Screen(i, j) = Tiles.GetTile(TileKey, (i Mod 8), (j Mod 8))
            
            'Debug.Print "i = " & i; "; j = " & j
        
        Next j
        
        Debug.Assert 1 = 1
        
    Next i
    
    Debug.Assert 1 = 1

End Sub

Private Sub ScreentoSheet(ws As Worksheet)
    Dim i As Long
    Dim j As Long
    Dim WriteColor As Long
    Dim WriteCell As range
    
    For i = LBound(Screen, 1) To UBound(Screen, 1)
        For j = LBound(Screen, 2) To UBound(Screen, 2)
            If Screen(i, j) Then
                WriteColor = TrueColor
            
            Else
                WriteColor = FalseColor
            
            End If
            
            ws.Cells((i + 1), (j + 1)).Interior.Color = WriteColor
        
        Next j
        
    Next i

End Sub

Private Sub InitalizeScreen(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = False
        
        Next j
    
    Next i

End Sub
