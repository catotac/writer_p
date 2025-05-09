Option Explicit

Private qManager As New QuestionnaireManager
Private btnDot(1 To 30, 1 To 5) As MSForms.CommandButton
Private lblQuestion(1 To 30) As MSForms.Label
Private btnReceive(1 To 30) As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Me.Caption = "Questionnaire - 30 Questions, 5 Columns"
    Me.Width = 750
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = 2200

    Dim colNames As Variant: colNames = Array("", "Column 1", "Column 2", "Column 3", "Column 4", "Column 5")

    Dim j As Integer
    For j = 1 To 5
        Dim lblHeader As MSForms.Label
        Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader" & j)
        With lblHeader
            .Caption = colNames(j)
            .Left = 300 + (j - 1) * 60
            .Top = 10
            .Width = 60
            .TextAlign = fmTextAlignCenter
            .Font.Bold = True
        End With
    Next j

    Dim i As Integer, topOffset As Integer
    For i = 1 To 30
        topOffset = 30 + (i - 1) * 35

        Set lblQuestion(i) = Me.Controls.Add("Forms.Label.1", "lblQ" & i)
        With lblQuestion(i)
            .Caption = i & ". " & qManager.Question(i)
            .Left = 20
            .Top = topOffset
            .Width = 260
        End With

        For j = 1 To 5
            Set btnDot(i, j) = Me.Controls.Add("Forms.CommandButton.1", "btnDot_" & i & "_" & j)
            With btnDot(i, j)
                .Caption = ""
                .Left = 300 + (j - 1) * 60 + 15
                .Top = topOffset
                .Width = 20
                .Height = 20
                .BackColor = RGB(220, 220, 220)
                .Tag = i & ":" & j
            End With
        Next j

        Set btnReceive(i) = Me.Controls.Add("Forms.CommandButton.1", "btnReceive" & i)
        With btnReceive(i)
            .Caption = "Receive"
            .Left = 620
            .Top = topOffset
            .Width = 60
            .Tag = i
        End With
    Next i

    ' Submit
    With cmdSubmit
        .Top = topOffset + 50
        .Left = 300
        .Width = 100
        .Caption = "Submit"
    End With

    ' Receive All
    With cmdReceiveAll
        .Top = topOffset + 50
        .Left = 420
        .Width = 100
        .Caption = "Receive All"
    End With
End Sub

Private Sub cmdSubmit_Click()
    qManager.SaveResponsesToDocument
    MsgBox "Responses recorded.", vbInformation
    Unload Me
End Sub

Private Sub cmdReceiveAll_Click()
    Dim i As Integer
    For i = 1 To 30
        SimulateAPIResponse i
        UpdateDotColors i
    Next i
End Sub

Private Sub btnReceive_Click()
    Dim ctrl As Control
    Set ctrl = Me.ActiveControl
    Dim i As Integer: i = CInt(ctrl.Tag)

    SimulateAPIResponse i
    UpdateDotColors i
End Sub

Private Sub btnDot_Click()
    Dim ctrl As Control
    Set ctrl = Me.ActiveControl

    Dim parts() As String: parts = Split(ctrl.Tag, ":")
    Dim i As Integer: i = CInt(parts(0))
    Dim j As Integer: j = CInt(parts(1))

    qManager.SelectedColumn(i) = j
    qManager.AnswerState(i) = "Yes"
    UpdateDotColors i
End Sub

Private Sub SimulateAPIResponse(i As Integer)
    Dim responseType As Integer: responseType = Int(Rnd() * 3) + 1
    Dim colChoice As Integer: colChoice = Int(Rnd() * 5) + 1

    Select Case responseType
        Case 1
            qManager.AnswerState(i) = "Yes"
            qManager.SelectedColumn(i) = colChoice
        Case 2
            qManager.AnswerState(i) = "No"
            qManager.SelectedColumn(i) = colChoice
        Case Else
            qManager.AnswerState(i) = "Not Present"
            qManager.SelectedColumn(i) = colChoice
    End Select
End Sub

Private Sub UpdateDotColors(i As Integer)
    Dim j As Integer
    For j = 1 To 5
        With btnDot(i, j)
            If qManager.SelectedColumn(i) = j Then
                Select Case qManager.AnswerState(i)
                    Case "Yes": .BackColor = RGB(0, 200, 0)
                    Case "No": .BackColor = RGB(200, 0, 0)
                    Case Else: .BackColor = RGB(160, 160, 160)
                End Select
            Else
                .BackColor = RGB(220, 220, 220)
            End If
        End With
    Next j
End Sub

