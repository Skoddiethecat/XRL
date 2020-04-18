Attribute VB_Name = "alFactory"
'@Folder("AltairLib")
Option Explicit

Private AltairLib As AltairLib

Public Function AltairLibLoad() As AltairLib
    If AltairLib Is Nothing Then Set AltairLib = New AltairLib
    
    Set AltairLibLoad = AltairLib

End Function

Public Sub AltairLibUnload()
    Set AltairLib = Nothing

End Sub
