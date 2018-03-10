Attribute VB_Name = "QuizModule"
Option Explicit


'##############################################################################
'# constants
'##############################################################################

Public Const ERROR_INVALID_DATA As Long = vbObjectError + 422
Public Const QUESTIONS_FILE As String = "questions.txt"


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
    Dim strQuestions As String: strQuestions = ReadQuestions()
        
    Dim strRec As Variant
    Set dictQ = New Scripting.Dictionary
    
    ' split notes area text in order to get an array containing each question data
    For Each strRec In Split(strQuestions, "####")
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
            objQRec.Notes = Trim(arrTmp(2))
            objQRec.Solution = Trim(arrTmp(3))
            objQRec.Points = Split(strId, "-")(1)
            dictQ.Add strId, objQRec
            
        End If
    Next
    
    CommonModule.Goto_Slide ("Board")
    
End Sub

'##############################################################################
'# Reads the questions
'##############################################################################
Function ReadQuestions() As String

    Dim strFilePath As String: strFilePath = ActivePresentation.Path + "\" + QUESTIONS_FILE
    If Len(Dir$(strFilePath)) = 0 Then
        MsgBox strFilePath & " file not found."
        Exit Function
    End If
    ReadQuestions = CommonModule.ReadTextFile(strFilePath)
    
    'the old version below read from notes area instead of external file:
    'ReadQuestions = Trim(CommonModule.Get_Slide("Board").NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text)

End Function

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
    With CommonModule.Get_Slide("QuestionSlide").Shapes("Question")
        .TextFrame2.TextRange.Text = dictQ(oShp.Name).Question
    End With
    With CommonModule.Get_Slide("QuestionSlide").Shapes("QuestionNote")
        .TextFrame2.TextRange.Text = dictQ(oShp.Name).Notes
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


