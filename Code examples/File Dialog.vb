Sub FileOpen()

'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .Filters.Add "All files", "*.*"
        'Show the dialog box
        .Show
        'Store in fullpath variable
        fullpath = .SelectedItems.Item(1)
    End With
    
    'It's a good idea to still check if the file type selected is accurate.
    'Quit the procedure if the user didn't select the type of file we need.
    If InStr(fullpath, ".xls") = 0 Then
        Exit Sub
    End If
 
    'Open the file selected by the user
    Workbooks.Open fullpath
    
End Sub