Attribute VB_Name = "CommonModule"
Option Explicit

Function Goto_Slide(SlideName As String)
    SlideShowWindows(1).View.GotoSlide (CommonModule.Get_Slide_Index(SlideName))
End Function

Function Get_Slide(SlideName As String) As Slide
    Set Get_Slide = ActivePresentation.Slides(Get_Slide_Index(SlideName))
End Function

Function Get_Slide_Index(SlideName As String)
    Dim sld As Slide
    For Each sld In ActivePresentation.Slides
        If sld.Name = SlideName Then
            Get_Slide_Index = sld.SlideIndex
            Exit Function
        End If
    Next sld
End Function
