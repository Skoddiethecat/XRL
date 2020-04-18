VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Classes")
Option Explicit

Public Tiles As Tiles
Public PlayField As PlayField

Private Sub Class_Initialize()
    Set Tiles = New Tiles
    Set PlayField = New PlayField

End Sub
