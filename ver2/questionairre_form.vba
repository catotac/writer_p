Option Explicit

Private qManager As New QuestionnaireManager
Private currentPage As Integer
Private comparisonLabels(1 To 30, 1 To 4) As MSForms.Label

Private Sub UserForm_Initialize()
    Me.Caption = "30 Questions Survey"
    currentPage = 1
    
    CreateQuestionsUI
    CreateComparisonUI
    ShowPage currentPage

    ' Sample Data for Testing
    Dim i As Integer, j As Integer
    For i = 1 To 30
        For j = 1 To 4
            If i Mod (j + 1) = 0 Then
                qManager.Response(i, j) = True
            ElseIf i Mod (j + 2) = 0 Then
                qManager.Response(i, j) = False
            End If
        Next j
    Next i

    UpdateComparisonUI
End Sub

' =========================
' Form Setup (Questionnaire UI)
' =========================

Private Sub CreateQuestionsUI()
    Dim i As Integer
    For i = 1 To 30
        With MultiPage1.Pages(0).Controls.Add("Forms.Label.1", "lblQuestion" & i)
            .Caption = i & ". " & qManager.Question(i)
            .Left = 30
            .Top = 20 + ((i - 1) Mod 10) * 40
            .Width = 380
            .Visible = False
        End With

        With MultiPage1.Pages(0).Controls.Add("Forms.OptionButton.1", "optYes" & i)
            .Caption = "Yes"
            .Left = 420
            .Top = 20 + ((i - 1) Mod 10) * 40
            .GroupName = "Group" & i
            .Visible = False
        End With

        With MultiPage1.Pages(0).Controls.Add("Forms.OptionButton.1", "optNo" & i)
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
        MultiPage1.Pages(0).Controls("lblQuestion" & i).Visible = show
        MultiPage1.Pages(0).Controls("optYes" & i).Visible = show
        MultiPage1.Pages(0).Controls("optNo" & i).Visible = show
    Next i
    cmdPrevious.Enabled = (pageNum > 1)
    cmdNext.Enabled = (pageNum < 3)
    currentPage = pageNum
End Sub

Private Sub SaveCurrentPageResponses()
    Dim i As Integer, startQ As Integer, endQ As Integer
    startQ = (currentPage - 1) * 10 + 1
    endQ = currentPage * 10

    For i = startQ To endQ
        If MultiPage1.Pages(0).Controls("optYes" & i).Value Then
            qManager.Response(i, 1) = True ' Only Resource 1 in form
        ElseIf MultiPage1.Pages(0).Controls("optNo" & i).Value Then
            qManager.Response(i, 1) = False
        Else
            qManager.Response(i, 1) = Empty
        End If
    Next i
End Sub

' =========================
' Form Setup (Comparison UI)
' =========================

Private Sub CreateComparisonUI()
    Dim i As Integer, j As Integer
    Dim leftBase As Integer: leftBase = 100
    Dim topBase As Integer: topBase = 30

    ' Headers
    For j = 1 To 4
        Dim lblHeader As MSForms.Label
        Set lblHeader = MultiPage1.Pages(1).Controls.Add("Forms.Label.1", "lblHeader" & j)
        With lblHeader
            .Caption = "Resource " & j
            .Left = leftBase + (j - 1) * 80
            .Top = 10
            .Width = 70
            .Height = 15
            .BackColor = RGB(220, 220, 220)
            .TextAlign = fmTextAlignCenter
        End With
    Next j

    ' Cells
    For i = 1 To 30
        For j = 1 To 4
            Set comparisonLabels(i, j) = MultiPage1.Pages(1).Controls.Add("Forms.Label.1", "cmp" & i & "_" & j)
            With comparisonLabels(i, j)
                .Left = leftBase + (j - 1) * 80
                .Top = topBase + (i - 1) * 18
                .Width = 70
                .Height = 16
                .BackStyle = fmBackStyleOpaque
                .TextAlign = fmTextAlignCenter
            End With
        Next j
    Next i
End Sub

Private Sub UpdateComparisonUI()
    Dim i As Integer, j As Integer
    For i = 1 To 30
        For j = 1 To 4
            Dim val As Variant
            val = qManager.Response(i, j)
            With comparisonLabels(i, j)
                If IsEmpty(val) Then
                    .Caption = "-"
                    .BackColor = RGB(180, 180, 180)
                ElseIf val = True Then
                    .Caption = "Yes"
                    .BackColor = RGB(0, 180, 0)
                Else
                    .Caption = "No"
                    .BackColor = RGB(200, 0, 0)
                End If
                .ForeColor = RGB(255, 255, 255)
            End With
        Next j
    Next i
End Sub

' =========================
' Event Handlers
' =========================

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
        If IsEmpty(qManager.Response(i, 1)) Then unanswered = unanswered + 1
    Next i

    If unanswered > 0 Then
        If MsgBox("You have " & unanswered & " unanswered questions. Submit anyway?", vbYesNo) = vbNo Then Exit Sub
    End If

    qManager.SaveResponsesToDocument
    Unload Me
    MsgBox "Thank you for completing the questionnaire!", vbInformation
End Sub

