
' This code creates a custom ribbon with a button that shows a form with 30 Yes/No questions
' To implement this in Word:
' 1. Press Alt+F11 to open the VBA editor
' 2. Insert a new module (Insert > Module)
' 3. Paste this code
' 4. Insert a new Class Module and name it "Ribbon"
' 5. Insert a new UserForm and name it "QuestionnaireForm"
' 6. Design the UserForm as described in the code
' 7. Create the callback as described in the CustomUI XML

' XML for the custom ribbon UI
Sub CreateRibbonXML()
    Dim xmlFile As String
    Dim fso As Object
    Dim xmlOutput As Object
    
    ' Path for the XML file
    xmlFile = Environ("TEMP") & "\customUI.xml"
    
    ' Create the XML content
    Dim xmlContent As String
    xmlContent = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
    "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf & _
    "  <ribbon>" & vbCrLf & _
    "    <tabs>" & vbCrLf & _
    "      <tab id=""customTab"" label=""Questionnaire"">" & vbCrLf & _
    "        <group id=""customGroup"" label=""Survey Form"">" & vbCrLf & _
    "          <button id=""customButton"" label=""Open Questionnaire"" " & _
    "imageMso=""QuestionMark"" size=""large"" onAction=""ShowQuestionnaire""/>" & vbCrLf & _
    "        </group>" & vbCrLf & _
    "      </tab>" & vbCrLf & _
    "    </tabs>" & vbCrLf & _
    "  </ribbon>" & vbCrLf & _
    "</customUI>"
    
    ' Create the file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create the XML file
    Set xmlOutput = fso.CreateTextFile(xmlFile, True)
    xmlOutput.Write xmlContent
    xmlOutput.Close
    
    MsgBox "XML file created at " & xmlFile & vbCrLf & vbCrLf & _
           "To use this ribbon:" & vbCrLf & _
           "1. Save this document as a .docm (macro-enabled document)" & vbCrLf & _
           "2. Close the document" & vbCrLf & _
           "3. Open the file in a ZIP program" & vbCrLf & _
           "4. Create a folder called 'customUI' if it doesn't exist" & vbCrLf & _
           "5. Add this XML file to that folder" & vbCrLf & _
           "6. Update the [Content_Types].xml file to include the relationship" & vbCrLf & _
           "7. Re-open the document", vbInformation, "Custom Ribbon Instructions"
End Sub

' Callback function for the ribbon button
Sub ShowQuestionnaire(control As IRibbonControl)
    QuestionnaireForm.Show
End Sub

' The following goes in the Ribbon class module
Option Explicit

Public ribbon As IRibbonUI

Sub OnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

' The following code goes in the code section of the QuestionnaireForm UserForm

' Array to store the questions
Dim Questions(1 To 30) As String
' Array to store responses (True = Yes, False = No, Empty = Not answered)
Dim Responses(1 To 30) As Variant
' Current page (each page shows 10 questions)
Dim currentPage As Integer

Private Sub UserForm_Initialize()
    ' Initialize questions
    Questions(1) = "Do you regularly back up your data?"
    Questions(2) = "Are you familiar with cloud storage solutions?"
    Questions(3) = "Have you experienced a computer virus in the past year?"
    Questions(4) = "Do you use a password manager?"
    Questions(5) = "Are your devices protected with a firewall?"
    Questions(6) = "Do you update your software regularly?"
    Questions(7) = "Have you received IT security training?"
    Questions(8) = "Do you use multi-factor authentication?"
    Questions(9) = "Are you concerned about data privacy?"
    Questions(10) = "Do you encrypt sensitive files?"
    Questions(11) = "Do you use public Wi-Fi networks?"
    Questions(12) = "Have you shared your password with colleagues?"
    Questions(13) = "Do you lock your computer when you leave your desk?"
    Questions(14) = "Are you aware of phishing techniques?"
    Questions(15) = "Have you clicked on suspicious email links recently?"
    Questions(16) = "Do you have antivirus software installed?"
    Questions(17) = "Is your operating system up to date?"
    Questions(18) = "Do you use unique passwords for different accounts?"
    Questions(19) = "Have you participated in security awareness programs?"
    Questions(20) = "Do you know your company's data security policies?"
    Questions(21) = "Have you used USB devices from unknown sources?"
    Questions(22) = "Do you report suspicious IT activities to your team?"
    Questions(23) = "Are you careful about what you download?"
    Questions(24) = "Do you use secure websites for online transactions?"
    Questions(25) = "Have you experienced data loss in the last year?"
    Questions(26) = "Do you know how to identify secure websites?"
    Questions(27) = "Are you careful about what you share on social media?"
    Questions(28) = "Do you know who to contact for IT security issues?"
    Questions(29) = "Have you read your company's IT usage policy?"
    Questions(30) = "Do you follow recommended security best practices?"
    
    ' Initialize the form
    Me.Caption = "30 Questions Survey"
    currentPage = 1
    
    ' Create UI elements dynamically
    CreateFormUI
    
    ' Load responses if they exist
    LoadResponses
    
    ' Show the first page
    ShowPage currentPage
End Sub

Private Sub CreateFormUI()
    ' Form dimensions
    Me.Width = 500
    Me.Height = 500
    
    ' Add page navigation and submit buttons
    With Me.Controls.Add("Forms.CommandButton.1", "cmdPrevious")
        .Caption = "Previous"
        .Width = 80
        .Height = 30
        .Left = 100
        .Top = 430
        .Visible = True
    End With
    
    With Me.Controls.Add("Forms.CommandButton.1", "cmdNext")
        .Caption = "Next"
        .Width = 80
        .Height = 30
        .Left = 200
        .Top = 430
        .Visible = True
    End With
    
    With Me.Controls.Add("Forms.CommandButton.1", "cmdSubmit")
        .Caption = "Submit"
        .Width = 80
        .Height = 30
        .Left = 300
        .Top = 430
        .Visible = True
    End With
    
    ' Add page indicators
    With Me.Controls.Add("Forms.Label.1", "lblPage")
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
            .Caption = i & ". " & Questions(i)
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
    Me.Controls("lblPage").Caption = "Page " & pageNum & " of 3"
    
    ' Show/hide questions based on current page
    Dim i As Integer
    For i = 1 To 30
        Dim show As Boolean
        show = (Int((i - 1) / 10) + 1 = pageNum)
        
        Me.Controls("lblQuestion" & i).Visible = show
        Me.Controls("optYes" & i).Visible = show
        Me.Controls("optNo" & i).Visible = show
    Next i
    
    ' Update button states
    Me.Controls("cmdPrevious").Enabled = (pageNum > 1)
    Me.Controls("cmdNext").Enabled = (pageNum < 3)
    
    ' Update current page
    currentPage = pageNum
End Sub

Private Sub cmdPrevious_Click()
    If currentPage > 1 Then
        SaveCurrentPageResponses
        ShowPage currentPage - 1
    End If
End Sub

Private Sub cmdNext_Click()
    If currentPage < 3 Then
        SaveCurrentPageResponses
        ShowPage currentPage + 1
    End If
End Sub

Private Sub SaveCurrentPageResponses()
    ' Save the current page's responses
    Dim startQ As Integer, endQ As Integer
    startQ = (currentPage - 1) * 10 + 1
    endQ = currentPage * 10
    
    Dim i As Integer
    For i = startQ To endQ
        If Me.Controls("optYes" & i).Value Then
            Responses(i) = True
        ElseIf Me.Controls("optNo" & i).Value Then
            Responses(i) = False
        Else
            Responses(i) = Empty ' Not answered
        End If
    Next i
End Sub

Private Sub LoadResponses()
    ' Load saved responses (if any)
    Dim i As Integer
    For i = 1 To 30
        If Not IsEmpty(Responses(i)) Then
            If Responses(i) = True Then
                Me.Controls("optYes" & i).Value = True
            Else
                Me.Controls("optNo" & i).Value = True
            End If
        End If
    Next i
End Sub

Private Sub cmdSubmit_Click()
    ' Save current page responses
    SaveCurrentPageResponses
    
    ' Check if all questions are answered
    Dim unanswered As Integer
    unanswered = 0
    
    Dim i As Integer
    For i = 1 To 30
        If IsEmpty(Responses(i)) Then
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
    SaveResponsesToDocument
    
    ' Close the form
    Unload Me
    
    ' Show completion message
    MsgBox "Thank you for completing the questionnaire!", vbInformation, "Submission Complete"
End Sub

Private Sub SaveResponsesToDocument()
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
        If IsEmpty(Responses(i)) Then
            response = "Not answered"
        ElseIf Responses(i) Then
            response = "Yes"
        Else
            response = "No"
        End If
        
        doc.Content.InsertAfter i & ". " & Questions(i) & ": " & response & vbCrLf
    Next i
    
    doc.Content.InsertAfter vbCrLf & "------------------------" & vbCrLf & vbCrLf
End Sub
