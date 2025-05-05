Option Explicit

Private m_Questions(1 To 30) As String
Private m_SelectedResponder(1 To 30) As Variant ' 1 to 4 or Empty

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

Public Property Get SelectedResponder(index As Integer) As Variant
    If index >= 1 And index <= 30 Then
        SelectedResponder = m_SelectedResponder(index)
    End If
End Property

Public Property Let SelectedResponder(index As Integer, value As Variant)
    If index >= 1 And index <= 30 Then
        m_SelectedResponder(index) = value
    End If
End Property

Public Sub SaveResponsesToDocument()
    Dim doc As Document
    Set doc = ActiveDocument

    doc.Content.InsertAfter vbCrLf & "QUESTIONNAIRE RESPONSES" & vbCrLf
    doc.Content.InsertAfter String(30, "=") & vbCrLf

    Dim i As Integer, responder As Variant
    For i = 1 To 30
        responder = m_SelectedResponder(i)
        If IsEmpty(responder) Then
            doc.Content.InsertAfter i & ". " & m_Questions(i) & ": Not answered" & vbCrLf
        Else
            doc.Content.InsertAfter i & ". " & m_Questions(i) & ": Answered by Responder " & responder & vbCrLf
        End If
    Next i
End Sub

