Attribute VB_Name = "GlobalConstants"
'@Folder("Main")
Option Explicit

Public Const ScreenResolutionColumns As Long = 640
Public Const ScreenResolutionRows As Long = 384

Public Const TrueColor As Long = 0
Public Const FalseColor As Long = 16777215

Sub printconsts()
    Debug.Print ScreenResolutionColumns
    Debug.Print ScreenResolutionRows
    Debug.Print TrueColor
    Debug.Print FalseColor

End Sub

