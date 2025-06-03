Sub ShowValidationPopup()
    ValidationForm.Show
End Sub

Function ToSymbol(val As Variant) As String
    Select Case LCase(Trim(val))
        Case "yes": ToSymbol = "✓"
        Case "no": ToSymbol = "✗"
        Case Else: ToSymbol = "☐"
    End Select
End Function

Sub PopulateValidationFormFromExcel(filePath As String, form As Object)
    Dim xlApp As Object, xlWB As Object, xlSheet As Object
    Dim data As Variant
    Dim lastRow As Long
    Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(filePath, False, True)
    Set xlSheet = xlWB.Sheets("ValidationData")

    ' Read header values
    form.txtCaseNumber.Text = xlSheet.Range("B1").Value
    form.txtCustomer.Text = xlSheet.Range("B2").Value

    ' Read tabular data
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    data = xlSheet.Range("A4:H" & lastRow).Value

    For i = 1 To UBound(data, 1)
        Dim t, q, src, intake, ecmp, letter, notes, callres As String
        t = Trim(data(i, 1))
        q = Trim(data(i, 2))
        src = ToSymbol(data(i, 3))
        intake = ToSymbol(data(i, 4))
        ecmp = ToSymbol(data(i, 5))
        letter = ToSymbol(data(i, 6))
        notes = data(i, 7)
        callres = data(i, 8)

        Dim prefix As String
        If t = "Complaint" Then
            prefix = "CQ" & Mid(q, 2)
        ElseIf t = "Taxonomy" Then
            prefix = "TQ" & Mid(q, 2)
        Else
            GoTo SkipRow
        End If

        On Error Resume Next
        form.Controls("lbl" & prefix & "Source").Caption = src
        form.Controls("lbl" & prefix & "Intake").Caption = intake
        form.Controls("lbl" & prefix & "ECMP").Caption = ecmp
        form.Controls("lbl" & prefix & "Letter").Caption = letter
        form.Controls("txt" & prefix & "Notes").Text = notes
        form.Controls("txt" & prefix & "Call").Text = callres
        On Error GoTo 0
SkipRow:
    Next i

    xlWB.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing
End Sub

