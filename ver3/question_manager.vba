Option Explicit

Private m_Questions(1 To 30) As String
Private m_SelectedColumn(1 To 30) As Variant
Private m_AnswerState(1 To 30) As String

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 1 To 30
        m_Questions(i) = "Question " & i & ": Sample question text?"
    Next i
End Sub

Public Property Get Question(index As Integer) As String
    Question = m_Questions(index)
End Property

Public Property Get SelectedColumn(index As Integer) As Variant
    SelectedColumn = m_SelectedColumn(index)
End Property

Public Property Let SelectedColumn(index As Integer, value As Variant)
    m_SelectedColumn(index) = value
End Property

Public Property Get AnswerState(index As Integer) As String
    AnswerState = m_AnswerState(index)
End Property

Public Property Let AnswerState(index As Integer, value As String)
    m_AnswerState(index) = value
End Property

Public Sub SaveResponsesToDocument()
    Dim doc As Document
    Set doc = ActiveDocument
    doc.Content.InsertAfter vbCrLf & "QUESTIONNAIRE RESPONSES" & vbCrLf
    doc.Content.InsertAfter String(30, "=") & vbCrLf

    Dim i As Integer
    For i = 1 To 30
        If IsEmpty(m_SelectedColumn(i)) Then
            doc.Content.InsertAfter i & ". " & m_Questions(i) & ": Not answered" & vbCrLf
        Else
            doc.Content.InsertAfter i & ". " & m_Questions(i) & ": Answered from Column " & m_SelectedColumn(i) & " [" & m_AnswerState(i) & "]" & vbCrLf
        End If
    Next i
End Sub

