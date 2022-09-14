VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private coll1 As New Collection
Private coll2 As New Collection
Private coll3 As New Collection
Private coll4 As New Collection

Public Sub Append(ByVal value1 As Variant, Optional ByVal value2 As Variant = Null, Optional ByVal value3 As Variant = Null, Optional ByVal value4 As Variant = Null)
    
    coll1.Add value1
    coll2.Add value2
    coll3.Add value3
    coll4.Add value4

End Sub

Public Sub Display()

    Dim strFilePath As String
    Dim ds As String
    Dim i As Long
    
    i = 1
    ds = Format(Now(), "ddmmmyyyy-hhmmss")
    ds = UCase(ds)
    strFilePath = ThisWorkbook.Path & "\SP Output " & ds & ".txt"
    
    Open strFilePath For Append As #1
    For i = 1 To coll1.Count
        Print #1, i & ":" & vbTab & coll1(i) & vbTab & coll2(i) & vbTab & coll3(i) & vbTab & coll4(i)
    Next i
    Close #1
    
    ThisWorkbook.Activate
    ActiveWorkbook.FollowHyperlink strFilePath

End Sub
