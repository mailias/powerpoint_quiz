Attribute VB_Name = "QuizModule"
Option Explicit


'##############################################################################
'# constants
'##############################################################################

Public Const ERROR_INVALID_DATA As Long = vbObjectError + 422


'##############################################################################
'# module variable
'##############################################################################

' a dictionary containing all questions
Public dictQ As Scripting.Dictionary

' stores the ID of the current question
Private vntCurQuestionId As String


'##############################################################################
'# start method: resets and starts the quiz
'##############################################################################
Sub Start()

    On Error GoTo -1

    Dim shape As shape
    
    ' -------------------------------
    ' reset shapes on board
    ' -------------------------------
    For Each shape In CommonModule.Get_Slide("Board").Shapes
        If InStr(1, shape.Name, "Q__") = 1 Then
            shape.Visible = msoTrue
            shape.ActionSettings(ppMouseClick).Run = "Click_Question_Button"
        End If
    Next
    
    ' -------------------------------
    ' reset shapes on point board
    ' -------------------------------
    With CommonModule.Get_Slide("PointBoard")
        .Shapes("Group1-Points").TextFrame.TextRange.Text = "0"
        .Shapes("Group2-Points").TextFrame.TextRange.Text = "0"
    End With
    
    ' -------------------------------
    ' reset questions slide
    ' -------------------------------
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("Question").TextFrame.TextRange.Text = "<Frage>"
        .Shapes("QuestionNote").TextFrame.TextRange.Text = "<Note>"
    End With
    
    ' -------------------------------
    ' read questions to dictionary
    ' -------------------------------
    Dim strNotesPageText As String
    Dim strRec As Variant
    Set dictQ = New Scripting.Dictionary
    strNotesPageText = Trim(CommonModule.Get_Slide("Board") _
                            .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text)
    ' split notes area text in order to get an array containing each question data
    For Each strRec In Split(strNotesPageText, "####")
        If strRec <> "" Then
            
            ' Debug.Print strRec
            Dim arrTmp() As String
            
            ' split the question in order to get its attributes
            arrTmp = Split(Trim(strRec), "---")
            If UBound(arrTmp) < 3 Or UBound(arrTmp) > 3 Then
                Dim strMsg As String: strMsg = arrTmp(0) & " hat zu viel oder zu wenig Wert-Trenner. " _
                    & "Erwartet ist folgendes Format ""ID --- Frage --- Antwortmöglichkeiten --- Lösung"". " _
                    & "Zeilenumbrüche und Leerzeilen sind erlaubt. "
                ' MsgBox strMsg
                Err.Raise ERROR_INVALID_DATA, "Fehler beim Einlesen der Fragen", strMsg
            End If
            
            Dim strId As String: strId = Trim_Improved(arrTmp(0))
            
            ' create question object and store in dictionary
            Dim objQRec As QuestionRecord: Set objQRec = New QuestionRecord
            objQRec.Id = strId
            objQRec.Question = Trim_Improved(arrTmp(1))
            objQRec.Notes = Trim_Improved(arrTmp(2))
            objQRec.Solution = Trim_Improved(arrTmp(3))
            objQRec.Points = Split(strId, "-")(1)
            dictQ.Add strId, objQRec
            
        End If
    Next
    
    CommonModule.Goto_Slide ("Board")
    
End Sub


'##############################################################################
'# a better trim that also removes linebreaks
'# @see CommonModule.Trim_Improved(String)
'##############################################################################
Function Trim_Improved(str As String)
    Trim_Improved = CommonModule.Trim_Improved(str)
End Function


'##############################################################################
'# this method is called when a question is selected on the game board
'# @param oShp the shape
'##############################################################################
Sub Click_Question_Button(oShp As shape)

    ' store the ID of the current question
    vntCurQuestionId = dictQ(oShp.Name).Id
       
    ' hide it on the board
    oShp.Visible = msoFalse
    
    ' set the question text and note/options on the question slide
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("Question").TextFrame.TextRange.Text = dictQ(oShp.Name).Question
        .Shapes("QuestionNote").TextFrame.TextRange.Text = dictQ(oShp.Name).Notes
    End With

    ' go to the question slide
    CommonModule.Goto_Slide ("QuestionSlide")

End Sub


'##############################################################################
'# this method is used in order to add points to the groups current score
'# @param oShp the shape
'##############################################################################
Sub Click_Plus(oShp As shape)

    ' get the score board
    Dim sld As Slide: Set sld = CommonModule.Get_Slide("PointBoard")
    
    ' select the score shape that belongs to the groups plus button
    Dim shpPoints As shape: Set shpPoints = sld.Shapes(Split(oShp.Name, "_")(1))
    
    ' increment the score by the question points
    shpPoints.TextFrame.TextRange.Text = shpPoints.TextFrame.TextRange.Text + dictQ(vntCurQuestionId).Points + 0
    
    ' set the current question marker to a negative value
    ' in order to prevent that points can be added to scores multiple times
    vntCurQuestionId = -1
    
End Sub


'##############################################################################
'# shows the solution
'# @param oShp the shape
'##############################################################################
Sub Click_Solution(oShp As shape)
    With CommonModule.Get_Slide("QuestionSlide")
        .Shapes("QuestionNote").TextFrame.TextRange.Text = dictQ(vntCurQuestionId).Solution
    End With
End Sub


