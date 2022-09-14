Attribute VB_Name = "ToolsMod"
Option Explicit

''' Tools Module
''' David Leinen JULY 2022
''' daleinen@quad.com
''' daleinen@hotmail.com

''' *** This module contains various tools, similar from project to project, but can/should be ***

''' NAME:               DESCRIPTION:
''' ActivateBook        SUBROUTINE TO ACTIVATE A CERTAIN OPEN WORKBOOK BY PARTIAL NAME
''' AddSubCols          FUNCTION TO ADD/SUBTRACT COLUMNS TO CELL ADDRESS AS STRING
''' AddSubRows          FUNCTION TO ADD/SUBTRACT ROWS TO CELL ADDRESS AS STRING
''' ArrayDims           FUNCTION TO RETURN THE NUMBER OF DIMENSIONS OF AN ARRAY
''' ArrayLen            FUNCTION TO RETURN LENGTH OF ARRAY
''' ArrayMatch          FUNCTION TO RETURN THE POSITION OF LOOK UP VALUE(FIRST OCCURNACE)
''' ArrayToString       FUNCTION TO CONVERT AN ARRAY OF VARIANT TO STRING
''' CaseClean(CC)       FUNCTION TO UPPER CASE AND REMOVE UNNEEDED SPACES AND CHARS
''' CharCount           FUNCTION TO COUNT THE NUMBER OF INSTANCES OF CERTAIN CHAR
''' CheckFolder         SUBROUTINE TO CHECK IF FOLDER EXISTS, IF NOT CREATE IT
''' CloseBook           SUBROUTINE TO CLOSE WORKBOOK BY PARTIAL NAME
''' CloseWorkBooks      SUBROUTINE TO CLOSE ALL OPEN WORKBOOKS EXCEPT THE ACTIVE ONE
''' ConvertDate         FUNCTION TO CONVERT VARIANT VALUE TO DATE
''' ConvertDouble       FUNCTION TO CONVERT VARIANT VALUE TO DOUBLE
''' ConvertLong         FUNCTION TO CONVERT VARIANT VALUE TO LONG
''' CopyCollection      FUNCTION TO COPY ONE COLLECTION TO ANOTHER
''' EndProgram          SUBROUTINE TO END A PROGRAM DISPLAYING MESSAGE BOX
''' EndProgram2         SUBROUTINE TO END A PROGRAM DISPLAYING MESSAGE BOX BY MINS AND SECS
''' Errors              SUBROUTINE FOR ERROR HANDLING AT END OF PROGRAM
''' FilePicker          FUNCTION TO RETURN THE PATH OF FILE THAT USER PICKS
''' FindCell            FUNCTION TO RETURN THE CELL LOCATION OF A LOOKUP VALUE(FIRST OCCURRENCE)
''' FindCol             FUNCTION TO RETURN THE COL OF A LOOKUP VALUE(FIRST OCCURRENCE)
''' FindRow             FUNCTION TO RETURN THE ROW LOCATION OF A LOOKUP VALUE(FIRST OCCURRENCE)
''' FolderPicker        FUNCTION TO RETURN THE PATH OF A FOLDER USER PICKS
''' FormatPath          FUNCTION TO FORMAT A STRING PATH DOWN TO A CERTAIN CHAR LENGTH
''' FractionToDec       FUCNTION TO CONVERT A FRACTION TO A NUMERIC VALUE
''' GetColumnLet        FUNCTION TO RETURN THE COLUMN LETTER/S NAME
''' GetColumnNum        FUNCTION TO RETURN A COLUMNS NUMERIC VALUE BASED ON RANGE
''' GetFileName         FUNCTION TO RETURN A FULL FILE NAME BASED ON PORTION OF NAME
''' GetFilePath         FUNCTION TO RETURN A FULL PATH OF NAMED FILE
''' GetFirstBlankCol    FUNCTION TO RETURN THE FIRST CELL OF BLANK DATA BASED ON ROW NUMBER
''' GetFirstBlankRow    FUNCTION TO RETURN THE FIRST EMPTY CELL IN COLUMN
''' GetFirstCol         FUNCTION TO RETURN THE FIRST CELL WITH DATA WITHIN COLUMN, BASED ON ROW NUMBER
''' GetFirstRow         FUNCTION TO RETURN THE FIRST ROW OF DATA BASED ON COLUMN NUMBER
''' GetFullName         FUNCTION TO RETURN THE FULL NAME OF FILE BASED ON PART OF NAME
''' GetLastColLet       FUNCTION TO RETURN THE LAST COL OF SHEET BASED ON ROW NUMBER AS A LETTER
''' GetLastColNum       FUNCTION TO RETURN THE LAST COL OF SHEET BASED ON ROW NUMBER AS A NUMBER
''' GetLastRow          FUNCTION TO RETURN THE LAST ROW OF SHEET BASED ON COLUMN NUMBER
''' GetRandomNum        FUNCTION TO RETURN RANDOM NUMBER BETWEEN LNGMIN(LNGMIN) AND LNGMAX(MALNGMAX)
''' HasNumeric          FUNCTION TO CHECK A STRING VALUE FOR ANY NUMERIC CHARS
''' InsertString        FUNCTION TO INSERT A SUBSTRING INTO A STRING AT SPECIFIC CHAR
''' IsCanada            FUNCTION TO CHECK IF STATE IS CANADA PROVINCE AND RETURN TRUE IF SO
''' IsInArray           FUNCTION TO RETURN TRUE IF VALUE IS ON ARRAY, FALSE IF NOT FOUND
''' IsInColl            FUNCTION TO RETURN TRUE IF VALUE IS IN COLLECTION, FALSE IF NOT FOUND
''' IsLetter            FUNCTION TO RETURN TRUE IF STRVALUE CONTAINS ONLY APLHA LETTERS
''' IsNonNumeric        FUNCTION TO RETURN TRUE IF STRVALUE CONTAINS NON NUMERIC
''' IsReallyEmpty       FUNCTION TO RETURN TRUE IF VALUE CONTAINS NO DATA OR ONLY WHITESPACE
''' IsState             FUNCTION TO CHECK IF STATE IS US AND RETURN TRUE IF SO
''' NewWorkBook         SUBROUTINE TO CREATE A NEW WORKBOOK IN WORKING DIR
''' OpenAllFiles        SUBROUTINE TO LOOP THROUGH AND OPEN ALL FILES IN FOLDER
''' OpenFile            SUBROUTINE TO OPEN ONE FILE FROM A PATH
''' PERSpoilage         FUNCTION TO RETURN ADDED SPOLIAGE QTY VALUE PERIDOCIALS
''' PrintArray          SUBROUTINE TO LOOP THROUGH AND PRINT ALL ELEMENTS OF AN ARRAY
''' PrintColl           SUBROUTINE TO LOOP THROUGH AND PRINT ALL ELEMENTS OF A COLLECTION
''' PrintToFile         SUBROUTINE TO PRINT AND APPEND VALUE TO TEXT FILE
''' RemoveAfter         FUNCTION TO RETURN A SUBSTRING OF STRING BEFORE CERTAIN CHAR, FIRST OCCURNACE
''' RemoveBefore        FUNCTION TO RETURN A SUBSTRING OF STRING AFTER CERTAIN CHAR, LAST OCCURRENCE
''' RemoveBetween       FUNCTION TO REMOVE ALL CHARS BETWEEN TWO SPECIFIC CHARS
''' RemoveChar          FUNCTION TO REMOVE CERTAIN CHARS FROM STRING
''' RemoveFirstChar     FUNCTION TO REMOVE THE FIRST CHAR OF STRING
''' RemoveLastChar      FUNCTION TO REMOVE THE LAST CHAR OF STRING
''' ReplaceString       FUNCTION TO REPLACE CERTAIN CHARS OR SUBSTRINGS FROM STRING
''' RoundAllDown        FUNCTION TO ALWAYS ROUND A DOUBLE WITH DECIMAL DOWN
''' RoundAllUp          FUNCTION TO ALWAYS ROUND A DOUBLE WITH DECIMAL UP
''' STDSpoilage         FUNCTION TO RETURN ADDED SPOLIAGE QTY FOR STD/MARKETING MAIL
''' SaveFileAs          FUNCTION TO RETURN THE NAME OF A NEW USER CREATED FILE
''' SelectToArray       FUNCTION TO PLACE SELECTION INTO A 1 OR MULTI DIMENSIONAL ARRAY
''' SelectToCollect     FUNCTION TO CREATE A COLLECTION BASED ON A SELECTION RANGE
''' StripLetters        FUNCTION TO RETURN ONLY THE NUMBERS FROM A STRING
''' StripNumeric        FUNCTION TO RETURN ONLY THE LETTERS FROM A STRING
''' SuperMid            FUNCTION TO RETURN A PORTION OF STRING BETWEEN DELIMETERS
''' SuperTrim           FUNCTION TO REMOVE ALL EXTRA SPACES FROM STRING AND TRIM. OPTIONAL TO REMOVE ALL SPACES
''' SwapCells           SUBROUTINE TO SWAP THE VALUE OF 2 CELLS
''' UpperColumn         SUBROUTINE TO UPPERCASE ALL DATA WITHIN COLUMN
''' WasteSeconds        SUBROUTINE TO WAIT A CERTAIN AMOUNT OF SECONDS
''' ----------------------------------------------------------------------------------------------------------------------

Public Sub ActivateBook(ByVal strName As String)
' subroutine to activate a certain open workbook by partial name
' INPUT  -> strName, a string of partial name of workbook
' OUTPUT -> none
' NEEDED -> RemoveBefore

    Dim wb As Workbook
    
    strName = Trim$(RemoveBefore(strName, "\"))
    
    For Each wb In Application.Workbooks
        If InStr(1, UCase(CStr(wb.name)), UCase(strName), vbTextCompare) > 0 Then wb.Activate
    Next wb

End Sub

Public Function AddSubCols(ByVal strCell As String, ByVal lngAddSub As Long) As String
' function to add/subtract columns to cell address as string
' INPUT  -> strCell, a string of cell address
'        -> lngAddSub, a number long of columns to add or subtract
' OUTPUT -> return string of new cell address
' NEEDED -> StripLetters, GetColumnNum, StripNumeric

    Dim l As Variant
    Dim s As Long
    
    l = StripLetters(strCell)
    s = CLng(GetColumnNum(CStr(StripNumeric(strCell))))
    s = s + lngAddSub
    
    AddSubCols = GetColumnLet(s) + CStr(l)
    
End Function

Public Function AddSubRows(ByVal strCell As String, ByVal lngAddSub As Long) As String
' function to add/subtract rows to cell address as string
' INPUT  -> strCall, a string of cell address
'        -> lngAddSub, a number long of rows to add or subtract
' OUTPUT -> return string of new cell address
' NEEDED -> StripLetters, StripNumeric

    Dim l As Variant
    Dim s As String
    
    l = StripLetters(strCell)
    s = CStr(StripNumeric(strCell))
    l = l + lngAddSub
    
    AddSubRows = s + CStr(l)

End Function

Public Function ArrayDims(ByRef arr As Variant) As Long
' function to return the number of dimensions of an array
' INPUT  -> arr, an array variant to determine number of dimensions
' OUTPUT -> a long of number of dimensions of array
' NEEDED -> none

    Dim s As String
    Dim arDim As Byte
    Dim i As Integer: i = 0
    
    On Error Resume Next
    Do
        s = CStr(arr(0, i))
        If Err.Number > 0 Then
            arDim = i
            On Error GoTo 0
            Exit Do
        Else
             i = i + 1
        End If
    Loop
    
    If arDim = 0 Then arDim = 1
    ArrayDims = CLng(arDim)
    
End Function

Public Function ArrayLen(arr As Variant) As Integer
' function to return length of array
' INPUT  -> arr, an array of variant type to work with
' OUTPUT -> an integer of length of array
' NEEDED -> none

    ArrayLen = UBound(arr) - LBound(arr) + 1
    
End Function

Public Function ArrayMatch(ByVal strValue As String, arr() As String) As String
' function to return the position of look up value(first occurrence)
' INPUT  -> strValue, a value to check for in array
'        -> arr, an array to search within
' OUTPUT -> a string of array position e.g, (2:1)
' NEEDED -> none

    Dim i, j As Long
    Dim lngDims As Long: lngDims = ArrayDims(arr)

    For i = LBound(arr) To UBound(arr)
        If lngDims = 1 Then
            If strValue = arr(i) Then
                ArrayMatch = i
                Exit Function
            End If
        Else
            For j = 0 To lngDims - 1
                If strValue = arr(i, j) Then
                    ArrayMatch = i & ":" & j
                    Exit Function
                End If
            Next j
        End If
    Next i

    ArrayMatch = ""

End Function

Public Function ArrayToString(ByVal arrInput As Variant) As String()
' function to convert an array of variant to string
' INPUT  -> arrInput, an array of variant
' OUTPUT -> same array of type string
' TODO   -> code to work with 2d arrays
' NEEDED -> none

    Dim e As Variant
    Dim strArr() As String
    Dim n As Integer: n = 0

    If VarType(arrInput) <> 12 And VarType(arrInput) < 8000 Then
        ReDim strArr(1) As String
        strArr(0) = CStr(arrInput)
        ArrayToString = strArr
    Else
        ReDim strArr(UBound(arrInput)) As String
          For Each e In arrInput
            strArr(n) = CStr(e)
            n = n + 1
          Next
        ArrayToString = strArr
    End If
    
End Function

Public Function CC(ByVal value As Variant) As String
' function to upper case and remove unneeded spaces and chars
' INPUT  -> value, a variant value to work with
' OUTPUT -> a clean string value
' NEEDED -> SuperTrim

    Dim s As String
        
    s = CStr(value)
    s = SuperTrim(s)
    s = UCase(s)

    CC = s
    
End Function

Public Function CharCount(ByVal strValue As String, ByVal strChar As String) As Byte
' function to count the number of instances of certain char
' INPUT  -> strValue, the string we are searching within
'        -> strChar, the char we are looking for
' OUTPUT -> a byte of length of string
' NEEDED -> none

    CharCount = Len(strValue) - Len(Replace(strValue, strChar, ""))

End Function

Public Sub CheckFolder(ByVal strPathName As String, ByVal strFolderName As String)
' subroutine to check if folder exists, if not create it
' INPUT  -> strPathName, the path we are looking for folder in
'           strFolderName, the name of folder we are searching for
' OUTPUT -> none
' NEEDED -> none
 
    If Dir(strPathName & "\" & strFolderName, vbDirectory) <> vbNullString Then
    Else
        MkDir strPathName & "\" & strFolderName
    End If

End Sub

Public Sub CloseBook(ByVal strName As String)
' subroutine to close workbook by partial name
' INPUT  -> strName, a string of partial name of workbook
' OUTPUT -> none
' NEEDED -> RemoveBefore
    
    Dim wb As Workbook
        
        strName = Trim$(RemoveBefore(strName, "\"))
    
    For Each wb In Application.Workbooks
        If InStr(UCase(CStr(wb.name)), UCase(strName)) > 0 Then wb.Close SaveChanges:=False
    Next wb

End Sub

Public Sub CloseWorkbooks(Optional StrFile As String = "")
' subroutine to close all open workbooks except the active one
' INPUT  -> Optional file name of workbook not to close
' OUTPUT -> none
' NEEDED -> none
    
    Application.DisplayAlerts = False
    Dim mac As String: mac = ThisWorkbook.name
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name <> mac And wb.name <> StrFile Then
            wb.Close
        End If

    Next

End Sub

'Public Function ConvertDate(ByVal varValue As Variant) As Variant
'' function to convert a variant value to date
'' INPUT  -> varValue, an variant containing a date to convert
'' OUTPUT -> a variant of date, if non date returns N/A
'' NEEDED -> none
'
'    If IsDate(varValue) Or IsNumeric(varValue) Then
'        ConvertDate = CDate(varValue)
'    Else
'        ConvertDate = "N/A"
'    End If
'
'End Function

Public Function ConvertDouble(ByVal varValue As Variant) As Double
' function to convert a variant value to double
' INPUT  -> varValue, an variant containing a number to convert
' OUTPUT -> a long of value, if non numeric then returns zero
' NEEDED -> none

    If IsNumeric(varValue) Then
        ConvertDouble = CDbl(varValue)
    Else
        ConvertDouble = 0#
    End If
    
End Function

Public Function ConvertLong(ByVal varValue As Variant) As Long
' function to convert a variant value to long
' INPUT  -> varValue, an variant containing a number to convert
' OUTPUT -> a long of value, if non numeric then returns zero
' NEEDED -> none

    If IsNumeric(varValue) Then
        ConvertLong = CLng(varValue)
    Else
        ConvertLong = 0
    End If
    
End Function

Public Function CopyCollection(ByVal coll As Collection) As Collection
' function to copy one collection to another
' INPUT  -> coll, a collection to copy
' OUTPUT -> a copied collection
' NEEDED -> none

    Dim newColl As New Collection
    Dim i As Long
    
    For i = 1 To coll.Count
        newColl.Add coll(i)
    Next i
    
    Set CopyCollection = newColl
    
End Function

Public Sub EndProgram(ByVal dblStartTime As Double)
' subroutine to end a program displaying message box
' INPUT  -> dblStartTime, an double containing the start time of program
' OUTPUT -> none
' NEEDED -> none

    Dim d As Double
    d = Round(Timer - dblStartTime, 2)
    'd = d - 1 'to counter the waste seconds (1)
    MsgBox "COMPLETE: Program ran successfully in " & d & " seconds", vbInformation

    'Call Shell("explorer.exe" & " " & ThisWorkbook.Path & "\Output", vbNormalFocus)

End Sub

Public Sub EndProgram2(ByVal sngTime As Single)
' subroutine to end a program displaying message box by mins and secs
' INPUT  -> sngTime, an double containing the start time of program
' OUTPUT -> none
' NEEDED -> none

    Dim d As Single
    Dim m As Byte, s As Byte
        
    d = Round(sngTime, 2)
        
    If d >= 60 Then
        m = WorksheetFunction.Floor(d / 60, 1)
        s = d Mod 60
        MsgBox "COMPLETE: Program ran successfully in " & m & " min " & s & " sec", vbInformation
    Else
        MsgBox "COMPLETE: Program ran successfully in " & Round(sngTime, 2) & " seconds", vbInformation
    End If

End Sub

Public Sub Errors(ByVal strDes As String, ByVal intNum As Integer)
' subroutine for error handling at end of program
' INPUT  -> strDes, a description of the error
'           intNum, a number describing nature of error
' OUTPUT -> none
' NEEDED -> none
    
    myBar.EndProgressBar
    Application.ScreenUpdating = True
    Call MsgBox("The following error occurred: " & intNum & " " & strDes & _
    vbCrLf & vbCrLf & "Please check Tips/Common Mistakes and run again" & _
    vbCrLf & "If error persists contact CAM on record or submitter of job" & _
    vbCrLf & "You may also contact ceautomationteam@quad.com", vbExclamation)

End Sub

Public Function FilePicker(Optional ByVal strType As String = "") As String
' function to return the path of file that user picks
' INPUT  -> strType, indicates type of file where are looking to pick
'        -> PDF,P for .pdf
'        -> EXCEL,E for .xlsx, xls, xlsm
'        -> Default, no filter(any other string)
' OUTPUT -> a string name of file that user picks
' NEEDED -> none
    
    On Error Resume Next
    Dim fileName As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    strType = Trim(UCase(strType))
    
    With fd
        .Filters.Clear
        .AllowMultiSelect = False
        If strType = "PDF" Or strType = "P" Then
            .Filters.Add "PDF Files", "*.pdf; *.PDF", 1
            .Title = "Select an PDF File"
        ElseIf strType = "EXCEL" Or strType = "E" Then
            .Filters.Add "Excel Files", "*.xlsx; *.xls, *.xlsm, *.csv", 1
            .Title = "Select an Excel File"
        Else
            '.Filters.Clear
        End If
        .show
        fileName = .SelectedItems(1)
    End With
    
    'If CStr(filename) = "" Then End
    
    FilePicker = CStr(fileName)

End Function

Public Function FindCell(ByVal varValue As Variant, Optional ByVal bolExact As Boolean = False) As String
' function to return the cell location of a lookup value(first occurrence)
' INPUT  -> varValue, a variant value we are looking up, Optional bolExact for exact matches
' OUTPUT -> the cell address of varLookup
' NEEDED -> CC
' TODO   -> return all occurrences, not just the first

    Dim cell As Range
    
    If bolExact Then
        For Each cell In ActiveSheet.UsedRange
            If CC(cell.value) = CC(varValue) Then
                FindCell = cell.Address(False, False)
                Exit Function
            Else
                FindCell = "N/A"
            End If
        Next cell
    Else
        For Each cell In ActiveSheet.UsedRange
            If InStr(1, CStr(cell.value), CStr(varValue), vbTextCompare) > 0 Then
                FindCell = cell.Address(False, False)
                Exit Function
            Else
                FindCell = "N/A"
            End If
        Next cell
    End If

End Function

Public Function FindCol(ByVal strVal As String, Optional ByVal bolExact As Boolean = False) As String
' function to return the col location of a lookup value(first occurrence)
' INPUT  -> val, a string value we are looking up, Optional addsub value of long to add/subtract
' OUTPUT -> the column that val resides in
' NEEDED -> StripNumeric, AddSubCols

    Dim s As String
    
    s = strVal
    s = StripNumeric(FindCell(strVal, bolExact))
        
    FindCol = s
    
End Function

Public Function FindRow(ByVal strVal As String, Optional ByVal bolExact As Boolean = False) As Long
' function to return the row of a lookup value(first occurrence)
' INPUT  -> val, a string value we are looking up, Optional addsub value of long to add/subtract
' OUTPUT -> the row that val resides in
' NEEDED -> StripLetters
    
    If FindCell(strVal, bolExact) = "N/A" Then
        FindRow = 0
        Exit Function
    End If
    
    Dim l As Long
    l = StripLetters(FindCell(strVal, bolExact))
    FindRow = l
    
End Function

Public Function FolderPicker() As String
' function to return the path of a folder user picks
' INPUT  -> none
' OUTPUT -> the path of folder
' NEEDED -> none

    On Error Resume Next
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        .show
        sItem = .SelectedItems(1)
    End With
    
    'If CStr(sItem) = "" Then End
        
    FolderPicker = sItem

End Function

Public Function FormatPath(ByVal strPth As String, Optional ByVal lgnLen As Long = 64) As String
' function to format a string path down to a certain char length
' INPUT  -> strPath, the workign path, Optional lgnLen a length to truncate to
' OUTPUT -> returns truncated path
' NEEDED -> none

    Dim s As String, c As Long
    s = strPth
    c = 0
    
    While Len(s) > lgnLen
        c = InStr(1, s, "\", vbTextCompare)
        s = Right(s, Len(s) - c)
    Wend

    FormatPath = s
    
End Function

Public Function FractionToDec(ByVal strFraction As String) As Single
' function to convert a fraction to a numeric value
' INPUT  -> strFraction, format must be "#-#/#", "# #/#" or "##.###".
' OUTPUT -> returns the numeric equivalent of input.
' NEEDED -> none

    Dim i As Integer
    Dim n As Single, Num As Single, Den As Single
    
    strFraction = Trim$(strFraction)
    strFraction = Replace(strFraction, " ", "-")
    i = InStr(strFraction, "/")
    
    If i = 0 Then
        n = val(strFraction)
    Else
        Den = val(Mid$(strFraction, i + 1))
        If Den = 0 Then Error 11
            strFraction = Trim$(Left$(strFraction, i - 1))
            i = InStr(strFraction, "-")
        If i = 0 Then
            Num = val(strFraction)
        Else
            Num = val(Mid$(strFraction, i + 1))
            n = val(Left$(strFraction, i - 1))
        End If
    End If
    
    If Den <> 0 Then
        n = n + Num / Den
    End If
    
    FractionToDec = n

End Function

Public Function GetColumnLet(ByVal lngCol As Long) As String
' function to return a columns letter/s name by numeric value
' INPUT  -> intCol, an integer containing column number
' OUTPUT -> a string of the column as a letter/s
' NEEDED -> none
    
    Dim v As Variant
    v = Split(Cells(1, lngCol).Address(True, False), "$")
    GetColumnLet = v(0)

End Function

Public Function GetColumnNum(ByVal strCol As String) As Long
' function to return a columns numeric value based on range
' INPUT  -> intCol, a string of cell range
' OUTPUT -> a long of numeric column
' NEEDED -> HasNumeric
    
    If strCol = "N/A" Then
        GetColumnNum = 0
        Exit Function
    End If
    
    If HasNumeric(strCol) = False Then
        GetColumnNum = Range(strCol & "1").Columns.column
    Else
        GetColumnNum = Range(strCol).Columns.column
    End If
    
End Function

Public Function GetFileName(ByVal strFileName As String, ByVal strFilePath As String) As String
' function to return a full file name based on portion of name
' INPUT  -> strFileName, a string of name of file we are looking for
'        -> strFilePath, a string for the path we are looking in
' OUTPUT -> return string of full name of file
' NEEDED -> none

    Dim StrFile As String
        StrFile = Dir(strFilePath & "\*")

    Do While Len(StrFile) > 0
        If InStr(UCase(StrFile), UCase(strFileName)) > 0 Then GetFileName = StrFile
        StrFile = Dir
    Loop

End Function

Public Function GetFilePath(ByVal strFileName As String, ByVal strFilePath As String) As String
' function to return a full path of named file
' INPUT  -> strFileName, a string of name of file
'        -> strFilePath, a string for the path we are looking in
' OUTPUT -> return string of full path of file
' NEEDED -> none

    Dim StrFile As String
        StrFile = Dir(strFilePath & "\*")
    
    Do While Len(StrFile) > 0
        If InStr(StrFile, strFileName) > 0 Then GetFilePath = strFilePath & "\" & StrFile
        StrFile = Dir
    Loop

End Function

Public Function GetFirstBlankCol(ByVal lngRow As Long, Optional ByVal strStartCol As String = "A") As Long
' function to return the first cell of blank data based on row number
' INPUT  -> lngRow, a long of row to look in
' OUTPUT -> return byte of column first blank cell
' NEEDED -> GetLastColLet

    Dim r As Range
    For Each r In Range(strStartCol & lngRow, GetLastColLet(lngRow) & lngRow)
        If Len(r.value) = 0 Then
            GetFirstBlankCol = CLng(r.column)
            Exit For
        End If
    Next r

End Function

Public Function GetFirstBlankRow(ByVal strCol As String, ByVal lngStartRow As Long) As Long
' function to return the first empty cell in column
' INPUT  -> strCol, a string of column to look in
' OUTPUT -> return long of row first blank cell
' NEEDED -> GetLastColLet

    Dim r As Range
    For Each r In Range(strCol & lngStartRow, strCol & GetLastRow(GetColumnNum(strCol)) + 1)
        If Len(r.value) = 0 Then
            GetFirstBlankRow = CLng(r.row)
            Exit For
        End If
    Next r
    
End Function

Public Function GetFirstCol(ByVal lngRow As Long) As String
' function to return the first cell with data within column, based on row number
' INPUT  -> lngRow, a long of row to look in
' OUTPUT -> return string of column first blank cell
' NEEDED -> none

    Dim l As Long
    Dim r As Range
        
    For Each r In ActiveSheet.UsedRange.Rows(lngRow).Cells
        If Len(r.value) > 0 Then
            l = CLng(r.column)
            Exit For
        End If
    Next
    
    GetFirstCol = GetColumnLet(l)

End Function

Public Function GetFirstRow(ByVal lngCol As Long) As Long
' function to return the first row of blank data based on column number
' INPUT  -> lngCol, a column number we need first row of
' OUTPUT -> the first row containing data in column
' NEEDED ->

   Dim r As Range
   For Each r In ActiveSheet.UsedRange.Columns(lngCol).Cells
       If Len(r.value) > 0 Then
           GetFirstRow = r.row
           Exit For
       End If
   Next

End Function

Public Function GetFullName(ByVal strInput As String) As String
' function to return the full name of file based on part of name
' INPUT  -> strInput, a str of partial name of file
' OUTPUT -> a string of full name of file
' NEEDED -> none

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(UCase(wb.name), UCase(strInput)) > 0 Then GetFullName = wb.name
    Next wb

End Function

Public Function GetLastColLet(ByVal lngRow As Long) As String
' function to return the last col of sheet based on row number
' INPUT  -> lngRow, a row number we need last col of
' OUTPUT -> the last col containing data in row, in a letter format
' NEEDED -> none

    GetLastColLet = Cells(lngRow, Columns.Count).End(xlToLeft).column
    GetLastColLet = Split(Cells(1, CLng(GetLastColLet)).Address, "$")(1)
        
End Function

Public Function GetLastColNum(ByVal lngRow As Long) As Long
' function to return the last col of sheet based on row number
' INPUT  -> lngRow, a row number we need last col of
' OUTPUT -> the last col containing data in row, in a number format
' NEEDED -> none

    GetLastColNum = Cells(lngRow, Columns.Count).End(xlToLeft).column
        
End Function

Public Function GetLastRow(ByVal lngCol As Long) As Long
' function to return the last row of data based on column number
' INPUT  -> lngCol, a column number we need last row of
' OUTPUT -> the last row containing data in column
' NEEDED -> none

    GetLastRow = Cells(Rows.Count, lngCol).End(xlUp).row
    
End Function

Public Function GetRandomNum(ByVal lngMin As Long, ByVal lngMax As Long) As Long
' function to return random number between lngMin(lngMin) and lngMax(malngMax)
' INPUT  -> lngMin, a min number
'        -> lngMax, a max number
' OUTPUT -> a long of random between min and max(including min and max values)
' NEEDED -> none

    GetRandomNum = CLng((lngMax - lngMin + 1) * Rnd + lngMin)

End Function

Public Function HasNumeric(ByVal strValue As String) As Boolean
' function to check string value for any numeric chars
' INPUT  -> strValue, a string containing the data to review
' OUTPUT -> boolean answer TRUE if data is contains numeric chars, FALSE otherwise
' NEEDED -> none

    Dim i As Integer
    For i = 1 To Len(strValue)
        If IsNumeric(Mid(strValue, i, 1)) Then
            HasNumeric = True
            Exit Function
        End If
    Next i
    
    HasNumeric = False
     
End Function

Public Function InsertString(ByVal strInput As String, ByVal strAdd As String, ByVal lngPosition As Long) As String
' function to insert a substring into a string at specific char
' INPUT  -> strInput, a string of source value
'        -> strAdd, a sub string to insert
'        -> lngPosition, an int of specific char
' OUTPUT -> a string with substring inserted
' NEEDED -> none

    Dim b As String, a As String
    
    b = Left$(strInput, lngPosition - 1)
    a = Right$(strInput, Len(strInput) - Len(b))
    
    InsertString = b & strAdd & a

End Function

Public Function IsCanada(ByVal strState As String) As Boolean
' function to check if state is CAN province and return true if so
' INPUT  -> strState, a string containing the data to review, not necessarily a state
' OUTPUT -> boolean answer TRUE if data is CAN, FALSE otherwise
' NEEDED -> none

    Select Case Trim$(UCase(CStr(strState)))
        Case "AB", "BC", "MB", "NB", "NL", "NS", "ON"
            IsCanada = True
        Case "PE", "QC", "SK", "NT", "NU", "YT"
            IsCanada = True
        Case Else
            IsCanada = False
    End Select

End Function

Public Function IsInArray(ByVal varToBeFound As Variant, ByVal arrLookIn As Variant) As Boolean
' function to return TRUE if value is in array, false if not found
' INPUT  -> strToBeFound, a string containing the data to review
'           arrLookIn, an array to check for value in
' OUTPUT -> a boolean decision if item is in array
' NEEDED -> none
' TODO   -> better error checking on arr

    If VarType(arrLookIn) = 8 Then
        If arrLookIn = varToBeFound Then
            IsInArray = True
            Exit Function
        End If
        Exit Function
    End If
    
    Dim i
    For i = LBound(arrLookIn) To UBound(arrLookIn)
        If arrLookIn(i, 1) = varToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    
    IsInArray = False

End Function

Public Function IsInColl(ByVal varItem As Variant, ByVal coll As Collection) As Boolean
' function to return TRUE if value is in collection, false if not found
' INPUT  -> varItem, a string containing the data to review
'           coll, a collection to check for value in
' OUTPUT -> a boolean decision if item is in array
' NEEDED -> none

    Dim i As Long
    For i = 1 To coll.Count
        If CC(varItem) = CC(coll(i)) Then
            IsInColl = True
            Exit Function
        End If
    Next i
    
    IsInColl = False
    
End Function

Public Function IsLetter(ByVal strValue As String) As Boolean
' function to return TRUE if strValue contains only alpha letters
' INPUT  -> strValue, a string containing the data to review
' OUTPUT -> boolean value TRUE for all non numeric chars
' NEEDED -> none
        
        Dim i As Integer
        strValue = UCase(strValue)
    
    For i = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, i, 1))
            Case 65 To 90
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
    
End Function

Public Function IsNonNumeric(ByVal strValue As String) As Boolean
' function to return TRUE if strValue contains non numeric
' INPUT  -> strValue, a string containing the data to review
' OUTPUT -> boolean value TRUE for all non numeric chars
' NEEDED -> none

    Dim i As Integer
    For i = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, i, 1))
            Case 32 To 47, 58 To 126
                IsNonNumeric = True
            Case Else
                IsNonNumeric = False
                Exit For
        End Select
    Next
    
End Function

Public Function IsReallyEmpty(ByVal varValue As Variant) As Boolean
' function to return TRUE if value contains no data or only whitespace
' INPUT  -> varValue, a variant containing the data to review
' OUTPUT -> boolean value true or false
' NEEDED -> none

    Dim s As String
    s = CStr(varValue)
    s = Replace(s, " ", "")
    
    If Trim$(Len(s)) > 0 Then
        IsReallyEmpty = False
    Else
        IsReallyEmpty = True
    End If

End Function

Public Function IsState(ByVal strState As String) As Boolean
' function to check if state is US and return true if so
' INPUT  -> strState, a string containing the data to review, not necessarily a state
' OUTPUT -> boolean answer TRUE if data is STATE, FALSE otherwise
' NEEDED -> none

    Select Case Trim$(UCase(CStr(strState)))
        Case "AK", "AL", "AR", "AS", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "GU", _
                "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", _
                "MP", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH", "OK", "OR", _
                "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UM", "UT", "VA", "VI", "VT", "WA", "WI", _
                "WV", "WY"
            IsState = True
        Case Else
            IsState = False
    End Select

End Function

Public Sub NewWorkBook(strName As String, ByVal folder As String)
' subroutine to create a new workbook in working DIR
' INPUT  -> strName, a string name of new file
' OUTPUT -> none
' NEEDED -> none
' TODO   -> return name of workbook
    
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
    With NewBook
        .SaveAs fileName:=folder & "\" & strName & " " & RemoveChar(time(), ":") & ".xlsx"
    End With

End Sub

Public Sub OpenAllFiles(ByVal strFolderPath As String)
' subroutine to loop through and open all files in folder
' INPUT  -> strPath, a file path to the folder we are looking in
' OUTPUT -> none
' NEEDED -> none

    Dim MyFile As String
        MyFile = Dir(strFolderPath & "\*")
    
    Do While MyFile <> ""
        Workbooks.Open fileName:=strFolderPath & "\" & MyFile
        MyFile = Dir
    Loop

End Sub

Public Sub OpenFile(ByVal strPath As String)
' subroutine to open file from a full path
' INPUT  -> strPath, a path to file
' OUTPUT -> none
' NEEDED -> none
    
    Call Workbooks.Open(strPath).Activate
        
End Sub

Public Function PERSpoilage(ByVal lngValue As Long) As Long
' function to return added spoliage qty value PERIDOCIALS
' INPUT  -> lngValue, a long containing the data to review
' OUTPUT -> a data type long of added spoilage qty
' NEEDED -> RoundAllUp

    Dim y As Double
        y = CDbl(lngValue)
    
    Select Case y
    Case Is <= 4000
        y = 75
    Case Is <= 10000
        y = 150
    Case Is <= 25000
        y = (y * 0.01) + 100
    Case Is <= 50000
        y = (y * 0.0075) + 125
    Case Is <= 250000
        y = (y * 0.005) + 100
    Case Is <= 500000
        y = (y * 0.005)
    Case Is > 500000
        y = (y * 0.0025)
    Case Else
        y = 0
    End Select
    
    y = RoundAllUp(y)
    PERSpoilage = CLng(y)

End Function

Public Sub PrintArray(ByVal arrInput As Variant)
' subroutine to loop through and print all elements of an array
' INPUT  -> arr, an array we want to print values of
' OUTPUT -> none
' NEEDED -> ArrayDims
    
    Dim i, j As Long
    Dim lngDims As Long
        lngDims = ArrayDims(arrInput)
    
    If IsEmpty(arrInput) = True Then
        Debug.Print "PRINT ARRAY: arr is blank"
    ElseIf IsNumeric(arrInput) = True Then
        Debug.Print arrInput
    Else
        For i = LBound(arrInput) To UBound(arrInput)
            If lngDims = 1 Then
                Debug.Print arrInput(i)
            Else
                For j = 0 To lngDims - 1
                    Debug.Print arrInput(i, j),
                Next j
                Debug.Print
            End If
        Next i
    End If
    
End Sub

Public Sub PrintColl(ByVal coll As Collection)
' subroutine to loop through and print all elements of a collection
' INPUT  -> coll, an array we want to print values of
' OUTPUT -> none
' NEEDED -> none

    Dim i As Long
    For i = 1 To coll.Count
        Debug.Print coll(i)
    Next i
    
End Sub

Public Sub PrintToFile(ByVal varValue As Variant)
' subroutine to print/append value to text file
' INPUT  -> varValue, a variant value to append to txt file
' OUTPUT -> none
' NEEDED -> none
    
    Dim strFilePath As String
    Dim x As String
    x = CStr(varValue)
    strFilePath = ThisWorkbook.path & "\output.txt"
   
    Open strFilePath For Append As #1
    Write #1, x
    Close #1

End Sub

Public Function RemoveAfter(ByVal strInput As String, ByVal strRemove As String) As String
' function to return a substring of string before certain char, first occurrence
' INPUT  -> strInput, an input string to remove from
'        -> strRemove, a char to remove and everything following
' OUTPUT -> return substring of string after char
' NEEDED -> none

    If InStr(strInput, strRemove) > 0 Then
        RemoveAfter = Left$(strInput, InStr(1, strInput, strRemove) - 1)
    Else
        RemoveAfter = strInput
    End If

End Function

Public Function RemoveBefore(ByVal strInput As String, ByVal strRemove As String) As String
' function to return a substring of string after certain char, last occurrence
' INPUT  -> strInput, an input string to remove from
'        -> strRemove, a char to remove and everything preceding
' OUTPUT -> return substring of string before char
' NEEDED -> none

    While InStr(strInput, strRemove) > 0
        strInput = Right(strInput, Len(strInput) - InStr(strInput, strRemove) - (Len(strRemove) - 1))
    Wend
    
    RemoveBefore = strInput

End Function

Public Function RemoveBetween(ByVal StrData As String, ByVal s1 As String, ByVal s2 As String) As String
'function to remove all chars between two specific chars
' INPUT  -> strData, string data to work with
'        -> s1, s2 chars to remove everything in between s1 and s2
' OUTPUT -> return substring of strData with text removed
' NEEDED -> none

    Dim s As String
        s = StrData
    
    While InStr(s, s1) > 0 And InStr(s, s2) > InStr(s, s1)
        s = Left(s, InStr(s, s1) - 1) & Mid(s, InStr(s, s2) + 1)
    Wend
    
    RemoveBetween = Trim$(s)
    
End Function

Public Function RemoveChar(ByVal strInput As String, ByVal chrRemove As String) As String
' function to remove certain chars from string
' INPUT  -> strInput, a string we are searching within
'           chrRemove, a character to remove
' OUTPUT -> a clean string with char removed
' NEEDED -> none
    
    RemoveChar = Replace$(strInput, chrRemove, "")
        
End Function

Public Function RemoveFirstChar(ByVal strInput As String) As String
' function to remove the first char of string
' INPUT  -> strInput, a string we are working within
' OUTPUT -> a string with first char removed
' NEEDED -> none
    
    If Len(strInput) < 1 Then
        RemoveFirstChar = ""
    Else
        RemoveFirstChar = Right(strInput, Len(strInput) - 1)
    End If
    
End Function

Public Function RemoveLastChar(ByVal strInput As String) As String
' function to remove the last char of string
' INPUT  -> strInput, a string we are working within
' OUTPUT -> a string with last char removed
' NEEDED -> none
    
    If Len(strInput) < 1 Then
        RemoveLastChar = ""
    Else
        RemoveLastChar = Left(strInput, Len(strInput) - 1)
    End If
    
End Function

Public Function ReplaceString(ByVal strInput As String, ByVal strRemove As String, ByVal strReplace As String) As String
' function to replace certain chars or substrings from string
' INPUT  -> strInput, a string we are searching within
'           strRemove, a substring to remove
'           strReplace, a substring to replace
' OUTPUT -> a string with a substring replaced
' NEEDED -> none
    
    Dim s As String
        s = strInput
    
    While InStr(1, s, strRemove, vbTextCompare) > 0
        s = Replace$(s, strRemove, strReplace)
    Wend
    
    ReplaceString = s
    
End Function

Public Function RoundAllDown(ByVal dblValue As Double) As Long
' function to always round a double with decimal down
' INPUT  -> dblValue, a double value to round down
' OUTPUT -> a rounded down long
' NEEDED -> none

    Dim result As Double
    result = Round(dblValue)
    
    If result > dblValue Then
        RoundAllDown = result - 1
    Else
        RoundAllDown = result
    End If
    
End Function

Public Function RoundAllUp(ByVal dblValue As Double) As Long
' function to always round a double with decimal up
' INPUT  -> dblValue, a double value to round up
' OUTPUT -> a rounded up long
' NEEDED -> none

    Dim result As Double
    result = Round(dblValue)
    
    If result >= dblValue Then
        RoundAllUp = result
    Else
        RoundAllUp = result + 1
    End If
    
End Function

Public Function SaveFileAs(ByVal bolClose As Boolean) As String
' function to return the name of a new user created file
' INPUT  -> bolClose, a boolean if we are to close new file
' OUTPUT -> a string of new file name
' NEEDED -> none
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim strFileName As String, t As String
    Dim NewBook As Workbook
        
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs fileName:=strFileName
    If bolClose = True Then Workbooks(RemoveBefore(strFileName, "\")).Close SaveChanges:=True
    
    SaveFileAs = RemoveBefore(strFileName, "\")
    
End Function

Public Function SelectToArray() As String()
' function to place selection into a 1 or multi dimensional array of type string
' INPUT  -> none
' OUTPUT -> return an array string of 1 or multi dimensions
' NEEDED -> GetColumnNum, StripNumeric, StripLetters, RemoveBefore, RemoveAfter, RemoveChar, ArrayToString

    Dim arrSelection() As String
    Dim lngArrLen, lngStartRow, lngStartCol, i, j As Long
    Dim lngDim As Long: lngDim = GetColumnNum(StripNumeric(RemoveBefore(RemoveChar(Selection.Address, "$"), ":"))) - _
    GetColumnNum(StripNumeric(RemoveAfter(RemoveChar(Selection.Address, "$"), ":"))) + 1
    
    If lngDim = 1 Then
        SelectToArray = ArrayToString(Selection.value)
        Exit Function
    End If
    
    lngArrLen = StripLetters(RemoveBefore(Selection.Address, ":")) - StripLetters(RemoveAfter(Selection.Address, ":"))
    lngStartRow = CLng(Selection.row)
    lngStartCol = CLng(Selection.column)
    
    ReDim arrSelection(0 To lngArrLen, 0 To (lngDim - 1))

    For i = 0 To lngArrLen
        For j = 0 To (lngDim - 1)
            arrSelection(i, j) = CStr(Cells(lngStartRow + i, lngStartCol + j).value)
        Next
    Next
    
    SelectToArray = arrSelection
    
End Function

Public Function SelectToCollect() As Collection
' function to create a collection based on a selection range
' INPUT  -> none, selection from excel is input
' OUTPUT -> return a collection from a select
' NEEDED -> none

    Dim r As Range, c As Range
    Dim coll As Collection
    Set coll = New Collection
    Set r = Selection
    
    For Each c In r
        coll.Add c
    Next c
    
    Set SelectToCollect = coll
    
End Function

Public Function STDSpoilage(ByVal lngValue As Long) As Long
' function to return added spoliage qty for STD/MARKETING mail
' INPUT  -> lngValue, a long containing the data to review
' OUTPUT -> a data type long of added spoilage to orginal qty
' NEEDED -> RoundAllUp

    Dim y As Double
        y = CDbl(lngValue)
    
    Select Case y
    Case Is < 40000
        y = (y * 0.01) + 750
    Case Is <= 100000
        y = (y * 0.015) + 750
    Case Is <= 250000
        y = (y * 0.02)
    Case Is <= 500000
        y = (y * 0.015)
    Case Is > 500000
        y = (y * 0.01)
    Case Else
        y = 0
    End Select
    
    y = RoundAllUp(y)
    STDSpoilage = CLng(y)

End Function

Public Function StripLetters(ByVal strInput As String) As String
' function to return only the numbers from a string
' INPUT  -> strInput, a string containing the data to review
' OUTPUT -> a data type long, numbers from input sting
' NEEDED -> none
    
    If IsNonNumeric(strInput) = True Or strInput = "" Then
        StripLetters = ""
        Exit Function
    End If
    
    Dim s As String
    Dim i As Integer
    
    For i = 1 To Len(strInput)
        If Mid(strInput, i, 1) >= "0" And Mid(strInput, i, 1) <= "9" Then
            s = s + Mid(strInput, i, 1)
        End If
    Next
    
    StripLetters = CStr(s)

End Function

Public Function StripNumeric(ByVal strInput As String) As String
' function to return only the letters from a string
' INPUT  -> strInput, a string containing the data to review
' OUTPUT -> a data type string, letters from input sting
' NEEDED -> none
    
    Dim s As String
    Dim i As Integer
    
    For i = 1 To Len(strInput)
        If IsNumeric(Mid(strInput, i, 1)) = False Then
            s = s + Mid(strInput, i, 1)
        End If
    Next
    
    StripNumeric = CStr(s)

End Function

Public Function SuperMid(ByVal strValue As String, ByVal strDelimter As String, ByVal lngLocation As Long) As String
' function to return a portion of string between delimeters
' INPUT  -> strValue, a string value to search within
'        -> strDelimter, the delimter that seperates string
'        -> lngLocation, a long that gives delimeter location, starting with 1
' OUTPUT -> a string type at delimeter location
' NEEDED -> none

    Dim str As Long, stp As Long, ct As Long
    Dim counter As Long, c As String
    
    If Right(strValue, 1) <> strDelimter Then strValue = strValue & strDelimter
    
    For counter = 1 To Len(strValue)
        c = Mid(strValue, counter, 1)
        If c = strDelimter Then
            ct = ct + 1
            If ct = (lngLocation - 1) Then
                str = counter
            End If
        End If
        If ct = lngLocation Then
            stp = counter
            Exit For
        End If
    Next counter
    
    SuperMid = Mid(strValue, str + 1, stp - str - 1)
    
End Function

Public Function SuperTrim(ByVal value As String, Optional ByVal allspaces As Boolean = False) As String
' function to remove all extra spaces from string and trim. Optional to remove ALL spaces
' INPUT  -> value, a string value to trim
'        -> allspaces, a boolean to trim ALL spaces if True
' OUTPUT -> a string of no extra white spaces
' NEEDED -> none

    Dim t As String
    Dim c As String
    Dim i As Long
    
    t = value
        
    For i = 1 To Len(value)
        c = Asc(Mid(value, i, 1))
        If c < 32 Or c > 126 Then t = Replace(value, Chr(c), "")
    Next i
    
    If allspaces = True Then
        t = Replace(value, Chr(32), "")
    Else
        t = Trim$(t)
    End If
    
    SuperTrim = t

End Function

Public Sub SwapCells(ByVal cell1 As Variant, ByVal cell2 As Variant)
' subroutine to swap the value of 2 cells
' INPUT  -> cell1, a variant of cell 1 address
'        -> cell2, a variant of cell 2 address
' OUTPUT -> none
' NEEDED -> none

    Dim temp As Variant
    temp = Range(cell1).value
        
    Range(cell1).value = Range(cell2).value
    Range(cell2).value = temp

End Sub

Public Sub UpperColumn(ByVal strCol As String)
' subroutine to UPPERCASE all data within
' INPUT  -> strCol, a column of data to uppercase as a letter/s
' OUTPUT -> none
' NEEDED -> none

    With Range(strCol & "2", Cells(Rows.Count, strCol).End(xlUp))
        .value = Evaluate("INDEX(UPPER(" & .Address(External:=True) & "),)")
    End With
    
End Sub

Public Sub WasteSeconds(ByVal lngSec As Long)
' subroutine to wait certain amount of seconds
' INPUT  -> intSec, a long representing seconds to wait
' OUTPUT -> prints times to immediate window
' NEEDED -> none

    'Debug.Print "start wait:    " & Right(CStr(Now), 11)
    Application.Wait (Now + TimeValue("0:00:" & lngSec))
    'Debug.Print "end wait:      " & Right(CStr(Now), 11)
    
End Sub
