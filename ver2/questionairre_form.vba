Option Explicit

Private qManager As New QuestionnaireManager
Private optColumn(1 To 30, 1 To 4) As MSForms.OptionButton
Private lblQuestion(1 To 30) As MSForms.Label

Private Sub UserForm_Initialize()
    Me.Caption = "Questionnaire - Select Column for Each Question"
    Me.Width = 600
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = 1600
    Me.ScrollTop = 0

    ' Add column headers
    Dim colLabels(1 To 4) As MSForms.Label
    Dim colNames As Variant: colNames = Array("", "Column 1", "Column 2", "Column 3", "Column 4")
    Dim j As Integer
    For j = 1 To 4
        Set colLabels(j) = Me.Controls.Add("Forms.Label.1", "lblCol" & j)
        With colLabels(j)
            .Caption = colNames(j)
            .Left = 280 + (j - 1) * 60
            .Top = 10
            .Width = 60
            .TextAlign = fmTextAlignCenter
            .Font.Bold = True
        End With
    Next j

    ' Add question rows
    Dim i As Integer, topOffset As Integer
    For i = 1 To 30
        topOffset = 30 + (i - 1) * 35

        ' Question Label
        Set lblQuestion(i) = Me.Controls.Add("Forms.Label.1", "lblQ" & i)
        With lblQuestion(i)
            .Caption = i & ". " & qManager.Question(i)
            .Left = 20
            .Top = topOffset
            .Width = 250
            .Height = 18
        End With

        ' Option buttons for each column
        For j = 1 To 4
            Set optColumn(i, j) = Me.Controls.Add("Forms.OptionButton.1", "optCol_" & i & "_" & j)
            With optColumn(i, j)
                .Caption = ""
                .Left = 280 + (j - 1) * 60 + 20
                .Top = topOffset
                .Width = 20
                .GroupName = "Group_" & i ' ensures only 1 per question
            End With
        Next j
    Next i

    ' Submit Button
    With cmdSubmit
        .Top = topOffset + 50
        .Left = 250
        .Width = 100
        .Caption = "Submit"
    End With
End Sub

Private Sub cmdSubmit_Click()
    Dim i As Integer, j As Integer
    For i = 1 To 30
        For j = 1 To 4
            If optColumn(i, j).Value Then
                qManager.SelectedColumn(i) = j
                Exit For
            End If
        Next j
    Next i

    qManager.SaveResponsesToDocument
    Unload Me
    MsgBox "Thank you! Your responses have been recorded.", vbInformation
End Sub

