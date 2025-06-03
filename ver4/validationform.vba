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

    For i = 1 To rowCount
        Dim qID As String
        qID = IIf(sectionName = "Complaint", "CQ" & i, "TQ" & (i + 3))

        ' Row controls
        Set frm.Controls.Add("Forms.Label.1", "lbl" & qID).Caption = "Q" & Mid(qID, 3)
        Set frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Source")
        Set frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Intake")
        Set frm.Controls.Add("Forms.Label.1", "lbl" & qID & "ECMP")
        Set frm.Controls.Add("Forms.Label.1", "lbl" & qID & "Letter")
        Set frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Notes")
        Set frm.Controls.Add("Forms.TextBox.1", "txt" & qID & "Call")

        For j = 0 To 6
            With frm.Controls(frm.Controls.Count - (6 - j))
                .Left = leftPos + j * 100
                .Top = topPos + i * 25
                .Width = 90
                .Height = 18
            End With
        Next j
    Next i
End Sub

