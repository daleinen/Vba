VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private arr() As Variant
Private upper As Long

Private Sub Class_Initialize()
    
    Count = 0
    upper = 0

End Sub

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    
    If VarType(key) = 8 Then key = Trim$(key)
    If VarType(value) = 8 Then value = Trim$(value)
    
    If KeyInArray(key) = True Then Exit Sub
    
    ReDim Preserve arr(1, upper)
    arr(0, upper) = key
    arr(1, upper) = value
    upper = upper + 1
    
    Count = UBound(arr, 2) + 1
    
End Sub

Public Sub AddSubValue(ByVal key As Variant, ByVal value As Variant)

    If VarType(key) = 8 Then key = Trim$(key)
    If VarType(value) = 8 Then value = Trim$(value)
        
    If KeyInArray(key) = False Then
        Call Add(key, value)
        Exit Sub
    End If
    
    Dim i As Long: i = (upper - 1)
    For i = 0 To i
        If key = arr(0, i) Then
            If VarType(value) = 8 Then
                arr(1, i) = arr(1, i) & value
            Else
                arr(1, i) = arr(1, i) + value
            End If
            Exit For
        End If
    Next i

End Sub

Public Function AllKeys() As Collection

    Dim c As New Collection
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        c.Add arr(0, i)
    Next i
    
    Set AllKeys = c

End Function

Public Function AllValues() As Collection

    Dim c As New Collection
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        c.Add arr(1, i)
    Next i
    
    Set AllValues = c

End Function

Public Sub ChangeKey(ByVal key As Variant, ByVal newKey As Variant)
    
    If KeyInArray(key) = False Then
        Call Add(key, 0)
        Exit Sub
    End If
    
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            arr(0, i) = newKey
            Exit For
        End If
    Next i

End Sub

Public Sub ChangeValue(ByVal key As Variant, ByVal value As Variant)
    
    If KeyInArray(key) = False Then
        Call Add(key, value)
        Exit Sub
    End If
    
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            arr(1, i) = value
            Exit For
        End If
    Next i

End Sub

Public Sub Clear()
    
    upper = 0
    ReDim arr(1, upper)
    Count = 0
    
End Sub

Public Sub ClearValues()

    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        arr(1, i) = 0
    Next i

End Sub

Public Function Copy(ByRef scollOrg As SuperCollection) As SuperCollection
    
    Dim scoll As New SuperCollection
    Dim i As Long
    
    For i = 0 To scollOrg.Count - 1
        scoll.Add scollOrg.KeyIndex(i), scollOrg.ValueIndex(i)
    Next i
    
    Set Copy = scoll
     
End Function

Public Sub Delete(ByVal key As Variant)

    Dim i As Long, e As Long
    Dim n() As Variant
    
    If KeyInArray(key) = False Then Exit Sub
    If Count <= 1 Then
        Call Clear
        Exit Sub
    End If
    
    i = (upper - 1)
    ReDim n(1, upper - 2)

    For i = 0 To i
        If key = arr(0, i) Then
            e = i
            Exit For
        End If
    Next i

    For i = 0 To (upper - 2)
        If i < e Then
            n(0, i) = arr(0, i)
            n(1, i) = arr(1, i)
        Else
            n(0, i) = arr(0, i + 1)
            n(1, i) = arr(1, i + 1)
        End If
    Next i
    
    arr = n
    upper = upper - 1
    Count = Count - 1
    
End Sub

Public Sub Display()

    Dim i As Long: i = (upper - 1)
    Dim s1 As String, s2 As String
    
    If Count = 0 Then
        Debug.Print "---------- Super Collection is empty ----------"
        Exit Sub
    End If
    
    For i = 0 To i
        s1 = GetVarType(VarType(arr(0, i)))
        s2 = GetVarType(VarType(arr(1, i)))
        If Len(arr(0, i)) > 2 Then
            Debug.Print "KEY: " & arr(0, i) & " <" & s1 & ">" & _
            vbTab & vbTab & "VALUE: " & arr(1, i) & " <" & s2 & ">"
        Else
            Debug.Print "KEY: " & arr(0, i) & " <" & s1 & ">" & _
            vbTab & vbTab & vbTab & "VALUE: " & arr(1, i) & " <" & s2 & ">"
        End If
    Next i
    
    Debug.Print "-------------------------------------"
    
End Sub

Public Function GetIndex(ByVal key As Variant) As Long
    
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            GetIndex = i
            Exit For
        End If
    Next i

End Function

Public Function GetKeys(ByVal value As Variant) As Collection
    
    Dim c As New Collection
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If value = arr(1, i) Then
            c.Add arr(0, i)
        End If
    Next i
    
    Set GetKeys = c

End Function

Public Function GetValue(ByVal key As Variant) As Variant

    GetValue = 0
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            GetValue = arr(1, i)
            Exit For
        End If
    Next i

End Function

Public Function KeyIndex(ByVal index As Long) As Variant
    
    KeyIndex = arr(0, index)

End Function

Public Function KeyExists(ByVal key As Variant) As Boolean

    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            KeyExists = True
            Exit For
        Else
            KeyExists = False
        End If
    Next i

End Function

Public Sub Sort(Optional ByVal descending As Boolean = False)
    
    Dim arrN As Variant, arrO As Variant
    Dim coll As Collection
    Dim i As Long
    
    arrO = arr
    ReDim arrN(1, upper)
    Set coll = AllKeys
    
    If descending = False Then
        Call SortCollectionA(coll)
    Else
        Call SortCollectionD(coll)
    End If
    
    For i = 0 To (Count - 1)
        arrN(0, i) = coll(i + 1)
        arrN(1, i) = GetValue(coll(i + 1))
    Next i
    
    arr = arrN
    
End Sub

Public Function ValueIndex(ByVal index As Long) As Variant
    
    ValueIndex = arr(1, index)

End Function

Public Function ValueSet() As Collection
    
    Dim c As New Collection
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If IsInColl(arr(1, i), c) = False Then
            c.Add arr(1, i)
        End If
    Next i
    
    Set ValueSet = c
    
End Function

'| ---------------- PRIVATE SUBS/FUNCTIONS ---------------- |'

Private Function GetVarType(ByVal varValue As Variant) As String
    
    Dim b As Byte
    Dim s As String
    
    b = VarType(varValue)
    
    Select Case varValue
    Case Is = CInt(2)
        s = "int"
    Case Is = CInt(3)
        s = "lng"
    Case Is = CInt(4)
        s = "sgl"
    Case Is = CInt(5)
        s = "dbl"
    Case Is = CInt(6)
        s = "ccy"
    Case Is = CInt(7)
        s = "dte"
    Case Is = CInt(8)
        s = "str"
    Case Is = CInt(11)
        s = "bol"
    Case Is = CInt(14)
        s = "dec"
    Case Is = CInt(17)
        s = "byt"
    Case Is = CInt(20)
        s = "llg"
    Case Else
        s = "n/a"
    End Select
    
    GetVarType = s
    
End Function

Private Function IsInColl(ByVal varItem As Variant, ByVal coll As Collection) As Boolean

    Dim i As Long
    For i = 1 To coll.Count
        If varItem = coll(i) Then
            IsInColl = True
            Exit Function
        End If
    Next i
    
    IsInColl = False

End Function

Private Function KeyInArray(ByVal key As Variant) As Boolean
    
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If key = arr(0, i) Then
            KeyInArray = True
            Exit For
        Else
            KeyInArray = False
        End If
    Next i

End Function

Private Sub SortCollectionA(ByRef coll As Collection)

    Dim collN As New Collection
    Dim collS As New Collection
    Dim i As Long, j As Long
    Dim vTemp As Variant
    
    For i = 1 To coll.Count
        If WorksheetFunction.IsNumber(coll(i)) Then
            collN.Add coll(i)
        Else
            collS.Add coll(i)
        End If
    Next i
    
    For i = 1 To collN.Count - 1
        For j = i + 1 To collN.Count
            If CDbl(collN(i)) > CDbl(collN(j)) Then
                vTemp = collN(j)
                collN.Remove j
                collN.Add vTemp, CStr(vTemp), i
            End If
        Next j
    Next i

    For i = 1 To collS.Count - 1
        For j = i + 1 To collS.Count
            If UCase(CStr(collS(i))) > UCase(CStr(collS(j))) Then
                vTemp = collS(j)
                collS.Remove j
                collS.Add vTemp, CStr(vTemp), i
            End If
        Next j
    Next i
    
    Set coll = Nothing
    Set coll = New Collection
    
    For i = 1 To collN.Count
        coll.Add collN(i)
    Next i
    For i = 1 To collS.Count
        coll.Add collS(i)
    Next i

End Sub

Private Sub SortCollectionD(ByRef coll As Collection)

    Dim collN As New Collection
    Dim collS As New Collection
    Dim i As Long, j As Long
    Dim vTemp As Variant
    
    For i = 1 To coll.Count
        If WorksheetFunction.IsNumber(coll(i)) Then
            collN.Add coll(i)
        Else
            collS.Add coll(i)
        End If
    Next i
    
    For i = 1 To collN.Count - 1
        For j = i + 1 To collN.Count
            If CDbl(collN(i)) < CDbl(collN(j)) Then
                vTemp = collN(j)
                collN.Remove j
                collN.Add vTemp, CStr(vTemp), i
            End If
        Next j
    Next i

    For i = 1 To collS.Count - 1
        For j = i + 1 To collS.Count
            If UCase(CStr(collS(i))) < UCase(CStr(collS(j))) Then
                vTemp = collS(j)
                collS.Remove j
                collS.Add vTemp, CStr(vTemp), i
            End If
        Next j
    Next i
    
    Set coll = Nothing
    Set coll = New Collection
    
    For i = 1 To collS.Count
        coll.Add collS(i)
    Next i
    For i = 1 To collN.Count
        coll.Add collN(i)
    Next i

End Sub

