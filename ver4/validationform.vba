Private Sub UserForm_Initialize()
    ' Add Complaint Validation Table
    CreateValidationTable Me, "Complaint", 3, 20, 20

    ' Add Taxonomy Validation Table
    CreateValidationTable Me, "Taxonomy", 7, 20, 220
End Sub

Sub CreateValidationTable(frm As Object, sectionName As String, rowCount As Integer, leftPos As Integer, topPos As Integer)
    Dim frmSection As MSForms.Frame
    Set frmSection = frm.Controls.Add("Forms.Frame.1", "fra" & sectionName)
    
    With frmSection
        .Caption = sectionName & " Validation"
        .Left = leftPos
        .Top = topPos
        .Width = 700
        .Height = 30 + rowCount * 30
    End With

    ' Add Header Row
    Dim headers As Variant
    headers = Array("Description", "Source", "Intake", "ECMP", "Letter", "Pulse Notes", "Call Results")
    
    Dim i As Integer, j As Integer
    Dim lefts As Variant
    lefts = Array(10, 120, 190, 260, 330, 400, 520)

    For j = 0 To UBound(headers)
        Dim lbl As MSForms.Label
        Set lbl = frm.Controls.Add("Forms.Label.1", "lbl" & sectionName & "_H" & j)
        With lbl
            .Caption = headers(j)
            .Top = topPos + 5
            .Left = leftPos + lefts(j)
            .Width = 70
            .Height = 14
            .Font.Bold = True
        End With
    Next j

    ' Add Rows
    For i = 1 To rowCount
        Dim qID As String
        qID = IIf(sectionName = "Complaint", "CQ" & i, "TQ" & (i + 3)) ' Q1-Q3 for complaints, Q4+ for taxonomy

        ' Description label
        Dim lblDesc As MSForms.Label
        Set lblDesc = frm.Controls.Add("Forms.Label.1", "lbl" & qID)
        lblDesc.Caption = "Q" & IIf(sectionName = "Complaint", i, i + 3)
        lblDesc.Left = leftPos + lefts(0)
        lblDesc.Top = topPos + 25 + (i - 1) * 30
        lblDesc.Width = 50

        ' Symbol labels
        Dim col As Variant, ctrlName As String
        For j = 1 To 4
            ctrlName = "lbl" & qID & headers(j)
            Dim lblSym As MSForms.Label
            Set lblSym = frm.Controls.Add("Forms.Label.1", ctrlName)
            lblSym.Caption = "☐"
            lblSym.Left = leftPos + lefts(j)
            lblSym.Top = topPos + 25 + (i - 1) * 30
            lblSym.Width = 30
        Next j

        ' Notes TextBox
        Set txtNotes = frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Notes")
        With txtNotes
            .Left = leftPos + lefts(5)
            .Top = topPos + 25 + (i - 1) * 30
            .Width = 100
        End With

        ' Call Result TextBox
        Set txtCall = frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Call")
        With txtCall
            .Left = leftPos + lefts(6)
            .Top = topPos + 25 + (i - 1) * 30
            .Width = 100
        End With
    Next i
End Sub

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

