VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Classes.GameData")
Option Explicit

Private PlayFieldArray(0 To ((ScreenResolutionRows / 8) - 1), 0 To ((ScreenResolutionColumns / 8) - 1)) As String

Private Sub Class_Initialize()
    InitalizePlayfield arr:=PlayFieldArray

End Sub

Private Sub InitalizePlayfield(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    
    Dim k As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            'arr(i, j) = 32
            
            arr(i, j) = (32 + (k Mod 93))
            
            k = k + 1
            
        
        Next j
    
    Next i
    
    Debug.Assert 1 = 1

End Sub

Public Sub SetTile(ByVal value As String, ByVal row As Long, ByVal column As Long)
    PlayFieldArray(row, column) = value

End Sub

Public Function GetTile(ByVal row As Long, ByVal column As Long) As String
    GetTile = PlayFieldArray(row, column)

End Function
