Private Sub UserForm_Initialize()
    CreateValidationTable Me, "Complaint", 3, 20, 120
    CreateValidationTable Me, "Taxonomy", 7, 20, 250
End Sub

Private Sub cmdLoadExcel_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Excel File"
        .Filters.Add "Excel Files", "*.xlsx"
        If .Show = -1 Then
            Call PopulateValidationFormFromExcel(fd.SelectedItems(1), Me)
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Sub CreateValidationTable(frm As Object, sectionName As String, rowCount As Integer, leftPos As Integer, topPos As Integer)
    Dim i As Integer, j As Integer
    Dim headers As Variant
    headers = Array("Description", "Source", "Intake", "ECMP", "Letter", "Pulse Notes", "Call Results")

    ' Add table headers
    For j = 0 To 6
        Dim lblHeader As MSForms.Label
        Set lblHeader = frm.Controls.Add("Forms.Label.1", "lblH" & sectionName & j)
        With lblHeader
            .Caption = headers(j)
            .Left = leftPos + j * 100
            .Top = topPos
            .Width = 90
            .Font.Bold = True
        End With
    Next j

    ' Add rows
    For i = 1 To rowCount
        Dim qID As String
        qID = IIf(sectionName = "Complaint", "CQ" & i, "TQ" & (i + 3))

        Dim lblDesc As MSForms.Label
        Set lblDesc = frm.Controls.Add("Forms.Label.1", "lbl" & qID)
        lblDesc.Caption = "Q" & Mid(qID, 3)

        Dim lblSrc As MSForms.Label
        Set lblSrc = frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Source")

        Dim lblIntake As MSForms.Label
        Set lblIntake = frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Intake")

        Dim lblECMP As MSForms.Label
        Set lblECMP = frm.Controls.Add("Forms.Label.1", "lbl" & qID & "ECMP")

        Dim lblLetter As MSForms.Label
        Set lblLetter = frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Letter")

        Dim txtNotes As MSForms.TextBox
        Set txtNotes = frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Notes")

        Dim txtCall As MSForms.TextBox
        Set txtCall = frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Call")

        ' Positioning
        Dim rowTop As Integer
        rowTop = topPos + i * 25

        lblDesc.Move leftPos + 0, rowTop, 90, 18
        lblSrc.Move leftPos + 100, rowTop, 90, 18
        lblIntake.Move leftPos + 200, rowTop, 90, 18
        lblECMP.Move leftPos + 300, rowTop, 90, 18
        lblLetter.Move leftPos + 400, rowTop, 90, 18
        txtNotes.Move leftPos + 500, rowTop, 90, 18
        txtCall.Move leftPos + 600, rowTop, 90, 18
    Next i
End Sub

