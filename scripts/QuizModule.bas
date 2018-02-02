Attribute VB_Name = "QuizModule"
Option Explicit

Public Const ERROR_INVALID_DATA As Long = vbObjectError + 422

Public dictQ As Scripting.Dictionary

Private vntCurQuestionId As String




Sub Start()

    On Error GoTo -1

    Dim shape As shape
    
    ' reset shapes on board
    For Each shape In CommonModule.Get_Slide("Board").Shapes
        If InStr(1, shape.Name, "Q__") = 1 Then
            shape.Visible = msoTrue
            shape.ActionSettings(ppMouseClick).Run = "Click_Question_Button"
        End If
    Next
    
    ' reset shapes on point board
    With CommonModule.Get_Slide("PointBoard")
        .Shapes("Group1-Points").TextFrame.TextRange.Text = "0"
        .Shapes("Group2-Points").TextFrame.TextRange.Text = "0"
    End With
    
    ' reset questions slide
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("Question").TextFrame.TextRange.Text = "<Frage>"
        .Shapes("QuestionNote").TextFrame.TextRange.Text = "<Note>"
    End With
    
    ' read questions to dictionary
    Set dictQ = New Scripting.Dictionary
    Dim strNotesPageText As String
    strNotesPageText = Trim(CommonModule.Get_Slide("Board").NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text)
    Dim strRec As Variant
    For Each strRec In Split(strNotesPageText, "####")
        If strRec <> "" Then
            ' Debug.Print strRec
            Dim arrTmp() As String
            arrTmp = Split(Trim(strRec), "---")
            If UBound(arrTmp) < 3 Or UBound(arrTmp) > 3 Then
                Dim strMsg As String: strMsg = arrTmp(0) & " hat zu viel oder zu wenig Wert-Trenner. " _
                    & "Erwartet ist folgendes Format ""ID --- Frage --- Antwortmöglichkeiten --- Lösung"". " _
                    & "Zeilenumbrüche und Leerzeilen sind erlaubt. "
                ' MsgBox strMsg
                Err.Raise ERROR_INVALID_DATA, "Fehler beim Einlesen der Fragen", strMsg
            End If
            
            'Debug.Print "-------------------------------------"
            Dim strId As String: strId = Trim(Replace(Trim(arrTmp(0)), vbCr, ""))
            'Debug.Print "ID: """ & Id & """"
            'Debug.Print "Frage: " & Trim(arrTmp(1))
            'Debug.Print "Hinweis/Option: " & Trim(arrTmp(2))
            'Debug.Print "Antwort: " & Trim(arrTmp(3))
            Dim objQRec As QuestionRecord
            Set objQRec = New QuestionRecord
            objQRec.Id = strId
            objQRec.Question = Trim(arrTmp(1))
            objQRec.Notes = Trim(arrTmp(2))
            objQRec.Solution = Trim(arrTmp(3))
            objQRec.Points = Split(Trim(arrTmp(0)), "-")(1)
            dictQ.Add strId, objQRec
            
        End If
    Next
    
    'Dim key As Variant
    'For Each key In dictQ.Keys()
    '    Debug.Print """" & key & """ ---> " & dictQ(key).Id
    'Next key
    
    CommonModule.Goto_Slide ("Board")
    
End Sub



Sub Click_Question_Button(oShp As shape)

    'MsgBox oShp.Id & " -- " & oShp.name & " --- " & dictQ(oShp.name).Points
    vntCurQuestionId = dictQ(oShp.Name).Id
    'MsgBox vntCurQuestionId
    'MsgBox dictQ(vntCurQuestionId).Points
    
       
    oShp.Visible = msoFalse
    
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("Question").TextFrame.TextRange.Text = dictQ(oShp.Name).Question
        .Shapes("QuestionNote").TextFrame.TextRange.Text = dictQ(oShp.Name).Notes
    End With

    CommonModule.Goto_Slide ("QuestionSlide")

End Sub


Sub Click_Plus(oShp As shape)

    Dim sld As Slide: Set sld = CommonModule.Get_Slide("PointBoard")
    Dim shpPoints As shape: Set shpPoints = sld.Shapes(Split(oShp.Name, "_")(1))
    shpPoints.TextFrame.TextRange.Text = shpPoints.TextFrame.TextRange.Text + dictQ(vntCurQuestionId).Points + 0
    
End Sub

Sub Click_Solution(oShp As shape)
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("QuestionNote").TextFrame.TextRange.Text = dictQ(vntCurQuestionId).Solution
    End With
End Sub


