Attribute VB_Name = "mFiles"
Option Explicit
Public bError As Boolean

Function ExcelFileSelection() As String()
    Dim rngBrowse As Range, rngFolder As Range
    Dim xlFile As FileDialog
    Dim sFolder() As String
    Dim i As Long

    Set xlFile = Application.FileDialog(msoFileDialogFilePicker)
    
    bError = False
    
    With xlFile
        .Title = "Select your two files (hold CTRL to multi-select)."
        .AllowMultiSelect = True
        If .Show <> -1 Then GoTo NoSelection
        
        ReDim sFolder(0 To .SelectedItems.Count)
        
        sFolder(0) = .SelectedItems(1)
        sFolder(1) = .SelectedItems(2)
    End With
    
    If IsArray(sFolder) Then
        ExcelFileSelection = sFolder
    Else
        bError = True
    End If
    
    Exit Function
    
NoSelection:
    
    bError = True
End Function
