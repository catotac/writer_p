Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            txtFilePath.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub cmdGenerate_Click()
    If txtFilePath.Text = "" Then
        MsgBox "Please select an Excel file first.", vbExclamation
        Exit Sub
    End If

    ' Call macro logic
    Call InsertValidationTablesFromForm(txtFilePath.Text)

    ' Close form
    Unload Me
End Sub

