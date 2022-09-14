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

Dim arr() As Variant
Dim upper As Long

Private Sub Class_Initialize()

    upper = 0
    Count = 0

End Sub

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    
    If VarType(key) = 8 Then key = Trim$(key)
    If VarType(value) = 8 Then value = Trim$(value)
    
    If KeyInArray(key) = True Then
        Call MsgBox("ADD function" & vbCrLf & "KEY already exists! (" & key & ")", vbExclamation)
        Exit Sub
    End If
    
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

    Dim i As Long: i = (upper - 1)
    For i = 0 To i
        If key = arr(0, i) Then
            arr(0, i) = newKey
            Exit For
        End If
    Next i

End Sub

Public Sub ChangeValue(ByVal key As Variant, ByVal value As Variant)

    Dim i As Long: i = (upper - 1)
    For i = 0 To i
        If key = arr(0, i) Then
            arr(1, i) = value
            Exit For
        End If
    Next i

End Sub

Public Sub Delete(ByVal key As Variant)

    Dim i As Long, e As Long
    Dim n() As Variant
    
    If KeyInArray(key) = False Then
        Call MsgBox("DELETE function" & vbCrLf & "KEY does not exist! (" & key & ")", vbExclamation)
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
    For i = 0 To i
        Debug.Print "KEY: " & arr(0, i) & vbTab & vbTab & "VALUE: " & arr(1, i)
    Next i
    Debug.Print "------------------------------"
    
End Sub

Public Function GetKey(ByVal value As Variant) As Collection
    
    Dim c As New Collection
    Dim i As Long: i = (upper - 1)
    
    For i = 0 To i
        If value = arr(1, i) Then
            c.Add arr(0, i)
        End If
    Next i
    
    Set GetKey = c

End Function

Public Function GetValue(ByVal key As Variant) As Variant

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

Public Sub Sort()
    
    Dim arrN As Variant, arrO As Variant
    Dim coll As Collection
    Dim i As Long
    
    arrO = arr
    ReDim arrN(1, upper)
    
    Set coll = AllKeys
    Call SortCollectionX(coll)
    
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

Private Sub SortCollectionX(ByRef coll As Collection)

    Dim vItm As Variant, vTemp As Variant
    Dim i As Long, j As Long

    For i = 1 To coll.Count - 1
        For j = i + 1 To coll.Count
            If IsNumeric(coll(i)) Then
                If CDbl(coll(i)) > CDbl(coll(j)) Then
                    vTemp = coll(j)
                    coll.Remove j
                    coll.Add vTemp, CStr(vTemp), i
                End If
            Else
                If UCase(CStr(coll(i))) > UCase(CStr(coll(j))) Then
                    vTemp = coll(j)
                    coll.Remove j
                    coll.Add vTemp, CStr(vTemp), i
                End If
            End If
        Next j
    Next i

End Sub

