VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AltairLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AltairLib")
Option Explicit

Public alOptimizer As alOptimizer
Public alExporter As alExporter

Private Sub Class_Initialize()
    Set alOptimizer = New alOptimizer
    Set alExporter = New alExporter

End Sub
