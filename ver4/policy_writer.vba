Sub InsertValidationTables()
    Dim doc As Document
    Set doc = ActiveDocument

    ' Load Excel data
    Dim excelData As Variant
    excelData = LoadValidationDataFromExcel("C:\Path\To\Your\ValidationData.xlsx")

    ' Insert Header Table (2x2)
    Dim headerTbl As Table
    Set headerTbl = doc.Tables.Add(doc.Range, 2, 2)
    headerTbl.Cell(1, 1).Range.Text = "Cell Number"
    headerTbl.Cell(2, 1).Range.Text = "Customer or not?"
    headerTbl.Cell(1, 2).Merge headerTbl.Cell(2, 2)

    doc.Content.InsertParagraphAfter
    doc.Paragraphs.Last.Range.InsertAfter vbCrLf & "Complaint Validation" & vbCrLf

    ' Complaint Table
    InsertValidationSection doc, excelData, "Complaint", 5

    doc.Paragraphs.Last.Range.InsertAfter vbCrLf & "Taxonomy Validation" & vbCrLf

    ' Taxonomy Table
    InsertValidationSection doc, excelData, "Taxonomy", 12
End Sub

Function LoadValidationDataFromExcel(filePath As String) As Variant
    Dim xlApp As Object, xlWB As Object, xlSheet As Object
    Dim lastRow As Long
    Dim data As Variant

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Open(filePath, False, True)
    Set xlSheet = xlWB.Sheets("ValidationData")

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' xlUp
    data = xlSheet.Range("A2:I" & lastRow).Value ' 9 columns: A–I

    xlWB.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing

    LoadValidationDataFromExcel = data
End Function

Sub InsertValidationSection(doc As Document, data As Variant, sectionName As String, questionCount As Integer)
    Dim i As Integer, r As Integer
    Dim t As Table
    Set t = doc.Tables.Add(doc.Range(doc.Content.End - 1), questionCount + 2, 8)

    ' Top headers
    t.Cell(1, 1).Range.Text = "Column Validation"
    t.Cell(1, 2).Range.Text = "Source Result"
    t.Cell(1, 3).Range.Text = "Intake"
    t.Cell(1, 4).Range.Text = "ECMP"
    t.Cell(1, 5).Range.Text = "Letter"
    t.Cell(1, 6).Range.Text = "Notes"
    t.Cell(1, 7).Range.Text = "Results"
    t.Cell(1, 8).Range.Text = ""

    ' Split headers under first column
    t.Cell(2, 1).Range.Text = "Question"
    t.Cell(2, 2).Range.Text = "Description"

    ' Merge vertical headers
    t.Cell(1, 1).Merge t.Cell(2, 1)
    t.Cell(1, 2).Merge t.Cell(2, 2)

    ' Fill table content
    r = 3
    For i = 1 To UBound(data, 1)
        If data(i, 1) = sectionName Then
            t.Cell(r, 1).Range.Text = data(i, 2)
            t.Cell(r, 2).Range.Text = data(i, 3)
            t.Cell(r, 3).Range.Text = data(i, 4)
            t.Cell(r, 4).Range.Text = data(i, 5)
            t.Cell(r, 5).Range.Text = data(i, 6)
            t.Cell(r, 6).Range.Text = data(i, 7)
            t.Cell(r, 7).Range.Text = data(i, 8)
            t.Cell(r, 8).Range.Text = data(i, 9)
            r = r + 1
            If r > questionCount + 2 Then Exit For
        End If
    Next i
End Sub

