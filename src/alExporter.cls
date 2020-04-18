VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "alExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AltairLib.Modules")
Option Explicit

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    ThisWorkbook.Save
    
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim Dialog As Object
    Dim FSO As Object
    Dim VBComponent As Object
    Dim Directory As String
    Dim WorkbookName As String
    Dim Path As String
    Dim Extension As String
    
    Dim i As Long
    
    Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    i = 0
    
    WorkbookName = ThisWorkbook.Name
    WorkbookName = Left(WorkbookName, Len(WorkbookName) - 5)
    
    Directory = GetSetting(WorkbookName, "Source Control", "Repository Folder", ThisWorkbook.Path)
    
    Dialog.Title = "Select the source folder"
    
    Dialog.InitialFileName = Directory
        
    If Dialog.Show = -1 Then
        Directory = Dialog.SelectedItems(1) & "\"
        SaveSetting WorkbookName, "Source Control", "Repository Folder", Directory

    End If
    
    If Not FSO.FolderExists(Directory) Then
        FSO.CreateFolder Directory
        
    End If
    
    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        If VBComponent.Name <> "Sheet1" Then
            Select Case VBComponent.Type
                Case ClassModule, Document
                    Extension = ".cls"
                
                Case Form
                    Extension = ".frm"
                
                Case Module
                    Extension = ".bas"
                
                Case Else
                    Extension = ".txt"
                    
            End Select
                
                    
            On Error Resume Next
            Err.Clear
            
            Path = Directory & "\" & VBComponent.Name & Extension
            VBComponent.Export Path
            
            If Err.number <> 0 Then
                MsgBox "Failed to export " & VBComponent.Name & " to " & Path, vbCritical
            Else
                i = i + 1
                Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & Path
            End If
            
            On Error GoTo 0
        
        End If
        
        
    Next
    
End Sub

