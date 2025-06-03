Private Sub cmdLoadExcel_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Excel File"
        .Filters.Add "Excel Files", "*.xlsx"
        If .Show = -1 Then
            Call PopulateFromExcel(fd.SelectedItems(1))
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub PopulateFromExcel(filePath As String)
    Dim xlApp As Object, xlWB As Object, xlSheet As Object
    Dim lastRow As Long, data As Variant

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(filePath, False, True)
    Set xlSheet = xlWB.Sheets("ValidationData")

    txtCaseNumber.Text = xlSheet.Range("B1").Value
    txtCustomer.Text = xlSheet.Range("B2").Value

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    data = xlSheet.Range("A3:H" & lastRow).Value

    Dim i As Integer
    For i = 1 To UBound(data, 1)
        Dim t = Trim(data(i, 1))
        Dim q = Trim(data(i, 2))
        Dim src = data(i, 3)
        Dim intake = data(i, 4)
        Dim ecmp = data(i, 5)
        Dim letter = data(i, 6)
        Dim notes = data(i, 7)
        Dim callRes = data(i, 8)

        Dim prefix As String
        If t = "Complaint" Then
            prefix = "CQ" & Mid(q, 2)
        ElseIf t = "Taxonomy" Then
            prefix = "TQ" & Mid(q, 2)
        Else
            GoTo Skip
        End If

        Me.Controls("lbl" & prefix & "Src").Caption = ToSymbol(src)
        Me.Controls("lbl" & prefix & "Intake").Caption = ToSymbol(intake)
        Me.Controls("lbl" & prefix & "ECMP").Caption = ToSymbol(ecmp)
        Me.Controls("lbl" & prefix & "Letter").Caption = ToSymbol(letter)
        Me.Controls("txt" & prefix & "Notes").Text = notes
        Me.Controls("txt" & prefix & "Call").Text = callRes
Skip:
    Next i

    xlWB.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing
End Sub

Function ToSymbol(val As Variant) As String
    Select Case LCase(Trim(val))
        Case "yes": ToSymbol = "✓"
        Case "no": ToSymbol = "✗"
        Case Else: ToSymbol = ""
    End Select
End Function

