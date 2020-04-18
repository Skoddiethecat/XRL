Attribute VB_Name = "Factories"
'@Folder("Main")
Option Explicit

Private GameData As GameData
Private Renderer As Renderer

Public Function GameDataCreate() As GameData
    If GameData Is Nothing Then Set GameData = New GameData
    
    Set GameDataCreate = GameData

End Function

Public Sub GameDataDestroy()
    Set GameData = Nothing

End Sub

Public Function RendererCreate() As Renderer
    If Renderer Is Nothing Then Set Renderer = New Renderer
    
    Set RendererCreate = Renderer

End Function

Public Sub RendererDestroy()
    Set Renderer = Nothing

End Sub
