Option Explicit

Private qManager As New QuestionnaireManager
Private optResponder(1 To 30, 1 To 4) As MSForms.OptionButton
Private lblQuestion(1 To 30) As MSForms.Label

Private Sub UserForm_Initialize()
    Me.Caption = "Questionnaire - Select Answer Source"
    Me.Width = 600
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = 1400
    Me.ScrollTop = 0

    Dim i As Integer, j As Integer
    Dim topOffset As Integer

    For i = 1 To 30
        topOffset = 20 + (i - 1) * 35

        ' Question Label
        Set lblQuestion(i) = Me.Controls.Add("Forms.Label.1", "lblQ" & i)
        With lblQuestion(i)
            .Caption = i & ". " & qManager.Question(i)
            .Left = 20
            .Top = topOffset
            .Width = 280
            .Height = 18
        End With

        For j = 1 To 4
            ' Responder Option Button
            Set optResponder(i, j) = Me.Controls.Add("Forms.OptionButton.1", "optR_" & i & "_" & j)
            With optResponder(i, j)
                .Caption = "R" & j
                .Left = 320 + ((j - 1) * 50)
                .Top = topOffset
                .Width = 40
                .GroupName = "Group_" & i
            End With
        Next j
    Next i

    ' Submit Button
    With cmdSubmit
        .Top = topOffset + 40
        .Left = 250
        .Width = 100
        .Caption = "Submit"
    End With
End Sub

Private Sub cmdSubmit_Click()
    Dim i As Integer, j As Integer
    For i = 1 To 30
        For j = 1 To 4
            If optResponder(i, j).Value Then
                qManager.SelectedResponder(i) = j
                Exit For
            End If
        Next j
    Next i

    qManager.SaveResponsesToDocument
    Unload Me
    MsgBox "Thank you! Your responses have been recorded.", vbInformation
End Sub

