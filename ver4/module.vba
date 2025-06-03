Sub ShowValidationForm()
    ValidationForm.Show
End Sub

Sub FillValidationTablesFromExcel(filePath As String)
    Dim data As Variant
    data = LoadValidationDataFromExcel(filePath)

    Dim doc As Document
    Set doc = ActiveDocument

    Dim i As Integer, tbl As Table, rowOffset As Integer
    For i = 1 To UBound(data, 1)
        Dim tableType As String: tableType = data(i, 1)
        Dim rowIndex As Integer

        If tableType = "Complaint" Then
            Set tbl = doc.Tables(2) ' Complaint table
            rowOffset = 2
            rowIndex = GetRowIndex(tbl, data(i, 2), rowOffset)
        ElseIf tableType = "Taxonomy" Then
            Set tbl = doc.Tables(3) ' Taxonomy table
            rowOffset = 2
            rowIndex = GetRowIndex(tbl, data(i, 2), rowOffset)
        Else
            Continue For
        End If

        If rowIndex > 0 Then
            tbl.Cell(rowIndex, 3).Range.Text = SymbolFromText(data(i, 5)) ' Intake
            tbl.Cell(rowIndex, 4).Range.Text = SymbolFromText(data(i, 6)) ' ECMP
            tbl.Cell(rowIndex, 5).Range.Text = SymbolFromText(data(i, 7)) ' Letter
            tbl.Cell(rowIndex, 6).Range.Text = SymbolFromText(data(i, 8)) ' Notes
            tbl.Cell(rowIndex, 7).Range.Text = SymbolFromText(data(i, 9)) ' Results
        End If
    Next i
End Sub

Function SymbolFromText(text As String) As String
    Select Case LCase(Trim(text))
        Case "yes": SymbolFromText = "✓"
        Case "no": SymbolFromText = "✗"
        Case Else: SymbolFromText = "☐"
    End Select
End Function

Function GetRowIndex(tbl As Table, question As String, startRow As Integer) As Integer
    Dim r As Integer
    For r = startRow + 1 To tbl.Rows.Count
        If Trim(tbl.Cell(r, 1).Range.Text) Like "*" & question & "*" Then
            GetRowIndex = r
            Exit Function
        End If
    Next r
    GetRowIndex = -1
End Function

Function LoadValidationDataFromExcel(filePath As String) As Variant
    Dim xlApp As Object, xlWB As Object, xlSheet As Object
    Dim lastRow As Long, data As Variant

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(filePath, False, True)
    Set xlSheet = xlWB.Sheets("ValidationData")

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    data = xlSheet.Range("A2:I" & lastRow).Value

    xlWB.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing

    LoadValidationDataFromExcel = data
End Function

