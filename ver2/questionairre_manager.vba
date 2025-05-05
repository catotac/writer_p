Option Explicit

Private m_Questions(1 To 30) As String
Private m_Responses(1 To 30, 1 To 4) As Variant ' 4 responders

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 1 To 30
        m_Questions(i) = "Question " & i & ": Sample question text?"
    Next i
End Sub

Public Property Get Question(index As Integer) As String
    If index >= 1 And index <= 30 Then
        Question = m_Questions(index)
    End If
End Property

Public Property Get Response(questionIndex As Integer, responderIndex As Integer) As Variant
    If questionIndex >= 1 And questionIndex <= 30 And responderIndex >= 1 And responderIndex <= 4 Then
        Response = m_Responses(questionIndex, responderIndex)
    End If
End Property

Public Property Let Response(questionIndex As Integer, responderIndex As Integer, value As Variant)
    If questionIndex >= 1 And questionIndex <= 30 And responderIndex >= 1 And responderIndex <= 4 Then
        m_Responses(questionIndex, responderIndex) = value
    End If
End Property

Public Sub SaveResponsesToDocument()
    Dim doc As Document
    Set doc = ActiveDocument
    
    doc.Content.InsertAfter vbCrLf & "QUESTIONNAIRE RESPONSES" & vbCrLf
    doc.Content.InsertAfter String(30, "=") & vbCrLf

    Dim i As Integer, j As Integer, response As String
    For i = 1 To 30
        doc.Content.InsertAfter i & ". " & m_Questions(i) & vbCrLf
        For j = 1 To 4
            If IsEmpty(m_Responses(i, j)) Then
                response = "Not answered"
            ElseIf m_Responses(i, j) Then
                response = "Yes"
            Else
                response = "No"
            End If
            doc.Content.InsertAfter "  Resource " & j & ": " & response & vbCrLf
        Next j
    Next i
End Sub

