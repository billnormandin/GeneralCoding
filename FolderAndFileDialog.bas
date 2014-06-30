Public Function RC_GetFolder(strPath As String, Optional fFlag As Boolean = False) As String
' Function returns a Folder path as chosen by the user 
' unless the optional flag fFlag is true, in which case a specific file is returned
    
    Dim fldr As FileDialog
    Dim sItem As String
    
    If fFlag Then
    
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    
    Else
    
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    End If
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:

    RC_GetFolder = sItem
    Set fldr = Nothing
    
End Function
