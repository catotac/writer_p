' ===================================================
' STANDARD MODULE (Insert > Module)
' ===================================================
Option Explicit

' This is the macro that the user will run to show the questionnaire
Sub ShowQuestionnaire()
    ' Create and show the form
    Dim frm As New QuestionnaireForm
    frm.Show
End Sub

' ===================================================
' CLASS MODULE (Insert > Class Module, rename to "QuestionnaireManager")
' ===================================================
Option Explicit

' Array to store the questions
Private m_Questions(1 To 30) As String
' Array to store responses (True = Yes, False = No, Empty = Not answered)
Private m_Responses(1 To 30) As Variant

' Constructor
Private Sub Class_Initialize()
    ' Initialize questions
    m_Questions(1) = "Do you regularly back up your data?"
    m_Questions(2) = "Are you familiar with cloud storage solutions?"
    m_Questions(3) = "Have you experienced a computer virus in the past year?"
    m_Questions(4) = "Do you use a password manager?"
    m_Questions(5) = "Are your devices protected with a firewall?"
    m_Questions(6) = "Do you update your software regularly?"
    m_Questions(7) = "Have you received IT security training?"
    m_Questions(8) = "Do you use multi-factor authentication?"
    m_Questions(9) = "Are you concerned about data privacy?"
    m_Questions(10) = "Do you encrypt sensitive files?"
    m_Questions(11) = "Do you use public Wi-Fi networks?"
    m_Questions(12) = "Have you shared your password with colleagues?"
    m_Questions(13) = "Do you lock your computer when you leave your desk?"
    m_Questions(14) = "Are you aware of phishing techniques?"
    m_Questions(15) = "Have you clicked on suspicious email links recently?"
    m_Questions(16) = "Do you have antivirus software installed?"
    m_Questions(17) = "Is your operating system up to date?"
    m_Questions(18) = "Do you use unique passwords for different accounts?"
    m_Questions(19) = "Have you participated in security awareness programs?"
    m_Questions(20) = "Do you know your company's data security policies?"
    m_Questions(21) = "Have you used USB devices from unknown sources?"
    m_Questions(22) = "Do you report suspicious IT activities to your team?"
    m_Questions(23) = "Are you careful about what you download?"
    m_Questions(24) = "Do you use secure websites for online transactions?"
    m_Questions(25) = "Have you experienced data loss in the last year?"
    m_Questions(26) = "Do you know how to identify secure websites?"
    m_Questions(27) = "Are you careful about what you share on social media?"
    m_Questions(28) = "Do you know who to contact for IT security issues?"
    m_Questions(29) = "Have you read your company's IT usage policy?"
    m_Questions(30) = "Do you follow recommended security best practices?"
End Sub

' Property to get a question by index
Public Property Get Question(index As Integer) As String
    If index >= 1 And index <= 30 Then
        Question = m_Questions(index)
    Else
        Question = ""
    End If
End Property

' Property to get/set a response by index
Public Property Get Response(index As Integer) As Variant
    If index >= 1 And index <= 30 Then
        Response = m_Responses(index)
    End If
End Property

Public Property Let Response(index As Integer, value As Variant)
    If index >= 1 And index <= 30 Then
        m_Responses(index) = value
    End If
End Property

' Method to save responses to the document
Public Sub SaveResponsesToDocument()
    ' Save responses to the document
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Add a new section for responses if needed
    Dim foundSection As Boolean
    foundSection = False
    
    ' Check if responses section already exists
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If InStr(para.Range.Text, "QUESTIONNAIRE RESPONSES") > 0 Then
            foundSection = True
            Exit For
        End If
    Next para
    
    ' Add the section header if not found
    If Not foundSection Then
        doc.Content.InsertParagraphAfter
        doc.Content.InsertAfter "QUESTIONNAIRE RESPONSES" & vbCrLf
        doc.Content.InsertAfter "========================" & vbCrLf & vbCrLf
    End If
    
    ' Add timestamp
    doc.Content.InsertAfter "Responses submitted on: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf & vbCrLf
    
    ' Add responses
    Dim i As Integer
    For i = 1 To 30
        Dim response As String
        If IsEmpty(m_Responses(i)) Then
            response = "Not answered"
        ElseIf m_Responses(i) Then
            response = "Yes"
        Else
            response = "No"
        End If
        
        doc.Content.InsertAfter i & ". " & m_Questions(i) & ": " & response & vbCrLf
    Next i
    
    doc.Content.InsertAfter vbCrLf & "------------------------" & vbCrLf & vbCrLf
End Sub

' ===================================================
' USERFORM (Insert > UserForm, rename to "QuestionnaireForm")
' ===================================================
Option Explicit

' Reference to the questionnaire manager class
Private qManager As New QuestionnaireManager
' Current page (each page shows 10 questions)
Private currentPage As Integer

' Declare explicit variables for buttons - this helps with event handling
Private btnPrevious As MSForms.CommandButton
Private btnNext As MSForms.CommandButton
Private btnSubmit As MSForms.CommandButton
Private lblPageIndicator As MSForms.Label

Private Sub UserForm_Initialize()
    ' Initialize the form
    Me.Caption = "30 Questions Survey"
    currentPage = 1
    
    ' Create UI elements directly (not dynamically)
    CreateFormUI
    
    ' Show the first page
    ShowPage currentPage
End Sub

Private Sub CreateFormUI()
    ' Form dimensions
    Me.Width = 500
    Me.Height = 500
    
    ' Add page navigation and submit buttons - create them directly
    Set btnPrevious = Me.Controls.Add("Forms.CommandButton.1", "cmdPrevious")
    With btnPrevious
        .Caption = "Previous"
        .Width = 80
        .Height = 30
        .Left = 100
        .Top = 430
        .Visible = True
        .Enabled = False  ' Disabled initially (page 1)
    End With
    
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "cmdNext")
    With btnNext
        .Caption = "Next"
        .Width = 80
        .Height = 30
        .Left = 200
        .Top = 430
        .Visible = True
    End With
    
    Set btnSubmit = Me.Controls.Add("Forms.CommandButton.1", "cmdSubmit")
    With btnSubmit
        .Caption = "Submit"
        .Width = 80
        .Height = 30
        .Left = 300
        .Top = 430
        .Visible = True
    End With
    
    ' Add page indicators
    Set lblPageIndicator = Me.Controls.Add("Forms.Label.1", "lblPage")
    With lblPageIndicator
        .Caption = "Page 1 of 3"
        .Width = 100
        .Height = 20
        .Left = 200
        .Top = 410
        .Visible = True
    End With
    
    ' Create question labels and option buttons (10 per page)
    Dim i As Integer
    For i = 1 To 30
        ' Question label
        With Me.Controls.Add("Forms.Label.1", "lblQuestion" & i)
            .Caption = i & ". " & qManager.Question(i)
            .Width = 350
            .Height = 20
            .Left = 50
            .Top = 30 + ((i - 1) Mod 10) * 35
            .Visible = False
        End With
        
        ' Yes option
        With Me.Controls.Add("Forms.OptionButton.1", "optYes" & i)
            .Caption = "Yes"
            .Width = 50
            .Height = 20
            .Left = 400
            .Top = 30 + ((i - 1) Mod 10) * 35
            .GroupName = "Group" & i
            .Visible = False
        End With
        
        ' No option
        With Me.Controls.Add("Forms.OptionButton.1", "optNo" & i)
            .Caption = "No"
            .Width = 50
            .Height = 20
            .Left = 450
            .Top = 30 + ((i - 1) Mod 10) * 35
            .GroupName = "Group" & i
            .Visible = False
        End With
    Next i
End Sub

Private Sub ShowPage(pageNum As Integer)
    ' Update page indicator
    lblPageIndicator.Caption = "Page " & pageNum & " of 3"
    
    ' Show/hide questions based on current page
    Dim i As Integer
    For i = 1 To 30
        Dim show As Boolean
        show = (Int((i - 1) / 10) + 1 = pageNum)
        
        Me.Controls("lblQuestion" & i).Visible = show
        Me.Controls("optYes" & i).Visible = show
        Me.Controls("optNo" & i).Visible = show
        
        ' Also update the option buttons based on stored responses
        If Not IsEmpty(qManager.Response(i)) Then
            If qManager.Response(i) = True Then
                Me.Controls("optYes" & i).Value = True
            Else
                Me.Controls("optNo" & i).Value = True
            End If
        End If
    Next i
    
    ' Update button states
    btnPrevious.Enabled = (pageNum > 1)
    btnNext.Enabled = (pageNum < 3)
    
    ' Update current page
    currentPage = pageNum
End Sub

' Important: Use separate event handler subs for each button
Private Sub cmdPrevious_Click()
    ' This handles clicking the Previous button
    If currentPage > 1 Then
        SaveCurrentPageResponses
        ShowPage currentPage - 1
    End If
End Sub

Private Sub cmdNext_Click()
    ' This handles clicking the Next button
    If currentPage < 3 Then
        SaveCurrentPageResponses
        ShowPage currentPage + 1
    End If
End Sub

Private Sub cmdSubmit_Click()
    ' This handles clicking the Submit button
    ' Save current page responses
    SaveCurrentPageResponses
    
    ' Check if all questions are answered
    Dim unanswered As Integer
    unanswered = 0
    
    Dim i As Integer
    For i = 1 To 30
        If IsEmpty(qManager.Response(i)) Then
            unanswered = unanswered + 1
        End If
    Next i
    
    ' Notify user of unanswered questions
    If unanswered > 0 Then
        If MsgBox("You have " & unanswered & " unanswered questions. Do you want to submit anyway?", _
                  vbQuestion + vbYesNo, "Confirm Submission") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Store responses
    qManager.SaveResponsesToDocument
    
    ' Close the form
    Unload Me
    
    ' Show completion message
    MsgBox "Thank you for completing the questionnaire!", vbInformation, "Submission Complete"
End Sub

Private Sub SaveCurrentPageResponses()
    ' Save the current page's responses
    Dim startQ As Integer, endQ As Integer
    startQ = (currentPage - 1) * 10 + 1
    endQ = currentPage * 10
    
    Dim i As Integer
    For i = startQ To endQ
        If Me.Controls("optYes" & i).Value Then
            qManager.Response(i) = True
        ElseIf Me.Controls("optNo" & i).Value Then
            qManager.Response(i) = False
        Else
            qManager.Response(i) = Empty ' Not answered
        End If
    Next i
End Sub
