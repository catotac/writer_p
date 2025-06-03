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
        MsgBox "Please select an Excel file.", vbExclamation
        Exit Sub
    End If
    Call FillValidationTablesFromExcel(txtFilePath.Text)
    MsgBox "Tables filled successfully!", vbInformation
    Unload Me
End Sub

