Attribute VB_Name = "CommonModule"
Option Explicit


'##############################################################################
'# a better trim that also removes linebreaks
'# @param str the string to be trimmed
'# @return the trimmed string
'# @see: https://stackoverflow.com/questions/24048400/function-to-trim-leading-and-trailing-whitespace-in-vba
'##############################################################################
Function Trim_Improved(str As String)
    Dim RE As Object, ResultString As String
    Set RE = CreateObject("vbscript.regexp")
    RE.MultiLine = True
    RE.Global = True
    RE.Pattern = "^[\s\xA0]+|[\s\xA0]+$"
    Trim_Improved = RE.Replace(str, "")
End Function


'##############################################################################
'# navigates to the slide with the given name
'# @param SlideName the name of the slide (the slide.Name property)
'##############################################################################
Function Goto_Slide(SlideName As String)
    SlideShowWindows(1).View.GotoSlide (CommonModule.Get_Slide_Index(SlideName))
End Function


'##############################################################################
'# gets the slide with the given slide name
'# @param SlideName the name of the slide (the slide.Name property)
'# @return the slide
'##############################################################################
Function Get_Slide(SlideName As String) As Slide
    Set Get_Slide = ActivePresentation.Slides(Get_Slide_Index(SlideName))
    ' TODO add error handling if slide is not found
End Function


'##############################################################################
'# gets the slide index for a given slide name
'# @param SlideName the name of the slide (the slide.Name property)
'# @return the slides index
'##############################################################################
Function Get_Slide_Index(SlideName As String)
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        If sld.Name = SlideName Then
            Get_Slide_Index = sld.SlideIndex
            Exit Function
        End If
    Next sld
End Function
