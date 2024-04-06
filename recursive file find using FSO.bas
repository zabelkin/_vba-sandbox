Attribute VB_Name = "files_find"
Option Explicit


' just example of usage
Public Function test_run(appOriginal As Application, appHidden As Application) As String

    Dim sPathSelected As String
    Dim lFilesFound As Long: lFilesFound = 1
    Dim lTimer As Double

    'Call SysFunctionOff

    With appHidden.FileDialog(msoFileDialogFolderPicker) 'Запрашиваем целевую папку с пакетом
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = False
        .Title = "Выбор папки поиска"
        .ButtonName = "Выбрать папку"
        .Show
        If .SelectedItems.Count = 1 Then sPathSelected = .SelectedItems(1) Else: Exit Function
    End With

    appHidden.ThisWorkbook.Sheets(1).UsedRange.Offset(1).Clear
    lTimer = Timer
            
    Call reccursive_file_find(CreateObject("Scripting.FileSystemObject").GetFolder(sPathSelected), lFilesFound, appHidden)
    
    Debug.Print lFilesFound & " file(s) found"
    lTimer = Timer - lTimer
    
    test_run = lFilesFound & " files in " & CInt(lTimer) & " seconds, i.e. " & _
        Format(lFilesFound / lTimer, "0.0") & " file/s or " & _
        Format(lTimer / lFilesFound, "0.0") & " s/file"
    
    'Call SysFunctionOn
    
End Function


' the recursive function
Private Sub reccursive_file_find(folder_obj As Object, ByRef counter As Long, app As Application)

    Dim subfolder_obj As Object
    Dim file_obj As Object
    Dim wb As Workbook
    
    For Each file_obj In folder_obj.Files
        If file_obj.Name Like "*.xls*" Then
            counter = counter + 1
            app.ThisWorkbook.Sheets(1).Cells(counter, 1).Value = file_obj.ParentFolder.Path
            app.ThisWorkbook.Sheets(1).Cells(counter, 2).Value = file_obj.Name
            
            On Error GoTo failed
            Set wb = app.Workbooks.Open( _
                CreateObject("scripting.filesystemobject").GetFile(file_obj.Path).ShortPath, _
                ReadOnly:=True)
            ' Windows(wb.Name).WindowState = xlMinimized
            app.ThisWorkbook.Sheets(1).Cells(counter, 3).Value = wb.Sheets(1).Cells(1, 1).Value
            wb.Close savechanges:=False
failed:
            If Err.Number <> 0 Then
                app.ThisWorkbook.Sheets(1).Cells(counter, 3).Value = "не смогли открыть файл"
                app.ThisWorkbook.Sheets(1).Cells(counter, 3).Interior.Color = vbYellow
            End If
            On Error GoTo 0
            
        End If
    Next file_obj
    
    For Each subfolder_obj In folder_obj.Subfolders
        Call reccursive_file_find(subfolder_obj, counter, app)
    Next subfolder_obj
    
End Sub
