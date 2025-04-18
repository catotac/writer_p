Option Explicit

Private qManager As New QuestionnaireManager
Private currentPage As Integer

Private Sub UserForm_Initialize()
    Me.Caption = "30 Questions Survey"
    Me.Width = 600
    Me.Height = 500
    currentPage = 1
    CreateQuestionsUI
    ShowPage currentPage
End Sub

Private Sub CreateQuestionsUI()
    Dim i As Integer
    For i = 1 To 30
        With Me.Controls.Add("Forms.Label.1", "lblQuestion" & i)
            .Caption = i & ". " & qManager.Question(i)
            .Left = 30
            .Top = 20 + ((i - 1) Mod 10) * 40
            .Width = 380
            .Visible = False
        End With

        With Me.Controls.Add("Forms.OptionButton.1", "optYes" & i)
            .Caption = "Yes"
            .Left = 420
            .Top = 20 + ((i - 1) Mod 10) * 40
            .GroupName = "Group" & i
            .Visible = False
        End With

        With Me.Controls.Add("Forms.OptionButton.1", "optNo" & i)
            .Caption = "No"
            .Left = 480
            .Top = 20 + ((i - 1) Mod 10) * 40
            .GroupName = "Group" & i
            .Visible = False
        End With
    Next i
End Sub

Private Sub ShowPage(pageNum As Integer)
    lblPage.Caption = "Page " & pageNum & " of 3"
    Dim i As Integer
    For i = 1 To 30
        Dim show As Boolean
        show = (Int((i - 1) / 10) + 1 = pageNum)
        Me.Controls("lblQuestion" & i).Visible = show
        Me.Controls("optYes" & i).Visible = show
        Me.Controls("optNo" & i).Visible = show
        If Not IsEmpty(qManager.Response(i)) Then
            If qManager.Response(i) = True Then
                Me.Controls("optYes" & i).Value = True
            Else
                Me.Controls("optNo" & i).Value = True
            End If
        End If
    Next i
    cmdPrevious.Enabled = (pageNum > 1)
    cmdNext.Enabled = (pageNum < 3)
    currentPage = pageNum
End Sub

Private Sub cmdNext_Click()
    If currentPage < 3 Then
        SaveCurrentPageResponses
        ShowPage currentPage + 1
    End If
End Sub

Private Sub cmdPrevious_Click()
    If currentPage > 1 Then
        SaveCurrentPageResponses
        ShowPage currentPage - 1
    End If
End Sub

Private Sub cmdSubmit_Click()
    SaveCurrentPageResponses
    Dim unanswered As Integer, i As Integer
    For i = 1 To 30
        If IsEmpty(qManager.Response(i)) Then unanswered = unanswered + 1
    Next i

    If unanswered > 0 Then
        If MsgBox("You have " & unanswered & " unanswered questions. Submit anyway?", vbYesNo) = vbNo Then Exit Sub
    End If

    qManager.SaveResponsesToDocument
    Unload Me
    MsgBox "Thank you for completing the questionnaire!", vbInformation
End Sub

Private Sub SaveCurrentPageResponses()
    Dim i As Integer, startQ As Integer, endQ As Integer
    startQ = (currentPage - 1) * 10 + 1
    endQ = currentPage * 10

    For i = startQ To endQ
        If Me.Controls("optYes" & i).Value Then
            qManager.Response(i) = True
        ElseIf Me.Controls("optNo" & i).Value Then
            qManager.Response(i) = False
        Else
            qManager.Response(i) = Empty
        End If
    Next i
End Sub

