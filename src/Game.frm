VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Game 
   Caption         =   "XRL"
   ClientHeight    =   735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "Game.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Main.UI")
Option Explicit

Private AltairLib As AltairLib
Private GameData As GameData
Private Renderer As Renderer
Private GameInitalize As GameInitalize
Private Running As Boolean

Private StartText As String
Private HaltText As String

Private GameSheet As Worksheet

Private Sub UserForm_Initialize()
    Set AltairLib = alFactory.AltairLibLoad
    Set GameData = Factories.GameDataCreate
    Set Renderer = Factories.RendererCreate
    Set GameInitalize = New GameInitalize
    Running = False
    
    StartText = "Start Game"
    HaltText = "Save/Quit"
    
    Run.Caption = StartText

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Running Then
        Cancel = True
        MsgBox "Hey the game is running, fuck off!", vbCritical, "Fuck off!"
        Exit Sub
    
    End If
    
End Sub

Private Sub UserForm_Terminate()
    Set AltairLib = Nothing
    alFactory.AltairLibUnload
    
    Set GameData = Nothing
    Factories.GameDataDestroy
    
    Set Renderer = Nothing
    Factories.RendererDestroy

End Sub

Private Sub Run_Click()
    If Not Running Then
        Running = True
        GameInitalize.Start GameSheet:=GameSheet
        Run.Caption = HaltText
    
    Else
        Running = False
        GameInitalize.Halt GameSheet:=GameSheet
        Run.Caption = StartText
    
    End If
    
End Sub

'  \`*-.
'   )  _`-.
'  .  : `. .
'  : _   '  \
'  ; *` _.   `*-._
'  `-.-'          `-.
'    ;       `       `.
'    :.       .        \
'    . \  .   :   .-'   .
'    '  `+.;  ;  '      :
'    : '   |    ;       ;-.
'    ; '   : :`-:     _.`* ;
' .•' /  .•' ; .•`- +'  `*'
' `•-•   `•-•  `•-•'
