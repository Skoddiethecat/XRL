Attribute VB_Name = "DELETEME"
'@Folder("Main")
Option Explicit

Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (Err.number = 0)
    Err.Clear
End Function

Sub Test()
    Dim coll1 As New Collection
    coll1.Add Item:=Sheet1.range("A1"), Key:="1"
    coll1.Add Item:=Sheet1.range("A2"), Key:="2"
    Debug.Print HasKey(coll1, "1")

    Dim coll2 As New Collection
    coll2.Add Item:=1, Key:="1"
    coll2.Add Item:=2, Key:="2"
    Debug.Print HasKey(coll2, "A")
End Sub

Sub StupidTruncation()
    Dim MyNumber As Long
    
    MyNumber = Int(99.8)    ' Returns 99.
    

    MyNumber = Fix(99.2)    ' Returns 99.
    
    MyNumber = Int(-99.8)    ' Returns -100.
    MyNumber = Fix(-99.8)    ' Returns -99.
    
    MyNumber = Int(-99.2)    ' Returns -100.
    MyNumber = Fix(-99.2)    ' Returns -99.

End Sub

Sub inttest()
    Dim number As Long
    
    number = Int(5.6)
    
    Debug.Print number

End Sub
