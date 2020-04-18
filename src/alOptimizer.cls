VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "alOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AltairLib.Modules")
Option Explicit

Public Active As Boolean
Private PageBreaks As Boolean

Private Sub Class_Initialize()
    Active = False
    PageBreaks = True

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    If Active Then Disable
    
    On Error GoTo 0

End Sub

Public Sub Enable(Optional ByVal ScreenUpdating As Boolean = True, Optional ByVal DisplayStatusBar As Boolean = True, _
                  Optional ByVal Calculation As Boolean = True, Optional ByVal EnableEvents As Boolean = True, _
                  Optional ByVal PrintCommunication As Boolean = True, Optional ByVal DisplayPageBreaks As Boolean = True)
    With Application
        If ScreenUpdating Then .ScreenUpdating = False
        If DisplayStatusBar Then .DisplayStatusBar = False
        If Calculation Then .Calculation = xlCalculationManual
        If EnableEvents Then .EnableEvents = False
        If PrintCommunication Then .PrintCommunication = False
    
    End With
    
    If DisplayPageBreaks Then
        PageBreaks = ActiveSheet.DisplayPageBreaks
        
        If PageBreaks Then ActiveSheet.DisplayPageBreaks = False 'this is a sheet-level setting
        
    End If
    
    Active = True

End Sub

Public Sub Disable(Optional ByVal ScreenUpdating As Boolean = True, Optional ByVal DisplayStatusBar As Boolean = True, _
                  Optional ByVal Calculation As Boolean = True, Optional ByVal EnableEvents As Boolean = True, _
                  Optional ByVal PrintCommunication As Boolean = True, Optional ByVal DisplayPageBreaks As Boolean = True)
    With Application
        If ScreenUpdating Then .ScreenUpdating = True
        If DisplayStatusBar Then .DisplayStatusBar = True
        If Calculation Then .Calculation = xlCalculationAutomatic
        If EnableEvents Then .EnableEvents = True
        If PrintCommunication Then .PrintCommunication = True
    
    End With
    
    If DisplayPageBreaks Then
        If PageBreaks Then ActiveSheet.DisplayPageBreaks = True
    
    End If
    
    Active = False

End Sub

