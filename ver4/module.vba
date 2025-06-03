' Macro to open the Validation Viewer Form
Sub ShowValidationPopup()
    ValidationForm.Show
End Sub

' Function to convert Yes/No text to symbols
Function ToSymbol(val As Variant) As String
    Select Case LCase(Trim(val))
        Case "yes": ToSymbol = "✓"
        Case "no": ToSymbol = "✗"
        Case Else: ToSymbol = ""
    End Select
End Function

' Loads the Excel file and populates the form
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
        Dim tableType As String: tableType = Trim(data(i, 1))
        Dim desc As String: desc = Trim(data(i, 2))

        Dim src As String: src = ToSymbol(data(i, 3))
        Dim intake As String: intake = ToSymbol(data(i, 4))
        Dim ecmp As String: ecmp = ToSymbol(data(i, 5))
        Dim letter As String: letter = ToSymbol(data(i, 6))
        Dim notes As String: notes = data(i, 7)
        Dim callres As String: callres = data(i, 8)

        ' Determine control prefix: CQ for Complaint, TQ for Taxonomy
        Dim ctrlPrefix As String
        If tableType = "Complaint" Then
            ctrlPrefix = "CQ" & Mid(desc, 2)
        ElseIf tableType = "Taxonomy" Then
            ctrlPrefix = "TQ" & Mid(desc, 2)
        Else
            GoTo SkipRow
        End If

        ' Try to set controls using dynamic naming
        On Error Resume Next
        form.Controls("lbl" & ctrlPrefix & "Src").Caption = src
        form.Controls("lbl" & ctrlPrefix & "Intake").Caption = intake
        form.Controls("lbl" & ctrlPrefix & "ECMP").Caption = ecmp
        form.Controls("lbl" & ctrlPrefix & "Letter").Caption = letter
        form.Controls("txt" & ctrlPrefix & "Notes").Text = notes
        form.Controls("txt" & ctrlPrefix & "Call").Text = callres
        On Error GoTo 0

SkipRow:
    Next i

    xlWB.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing
End Sub

