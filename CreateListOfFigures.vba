Sub CreateListOfFigures()
    Dim slide As slide
    Dim shape As shape
    Dim figureCount As Integer
    Dim indexSlide As slide
    Dim yPos As Integer
    Dim maxFiguresPerSlide As Integer
    Dim currentFigures As Integer
            

    
    ' Define the maximum number of lines in the list of figures per slide
    maxFiguresPerSlideDef = 15 ' ### user input
    maxFiguresPerSlide = maxFiguresPerSlideDef
            
            
    ' Insert a new slide for the list of figures and choose the correct layer of the slide master
    Set indexSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutContentWithCaption) ' ### user input
    indexSlide.Shapes(1).TextFrame.TextRange.Text = "List of Figures"
    figureCount = 1
    currentFigures = 0
    
    ' Search for figures in every slide
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            If shape.Type = msoPicture Then
                ' Check if maximum number of figures per slide is reached
                If currentFigures >= maxFiguresPerSlide Then
                    ' Insert a new slide for the list of figures and choose the correct layer of the slide master
                    Set indexSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutContentWithCaption) ' ### user input
                    indexSlide.Shapes(1).TextFrame.TextRange.Text = "List of Figures"
                    maxFiguresPerSlide = maxFiguresPerSlide + maxFiguresPerSlideDef
                End If
                ' Add index number of figure and its alternative text as caption to the list of figures
                Dim altText As String
                altText = shape.AlternativeText
                If altText = "" Then
                    altText = "Kein Alternativtext"
                End If
                indexSlide.Shapes(2).TextFrame.TextRange.Text = indexSlide.Shapes(2).TextFrame.TextRange.Text & _
                    "Abbildung " & figureCount & ": " & altText & vbCrLf
                figureCount = figureCount + 1
                currentFigures = currentFigures + 1
            End If
        Next shape
    Next slide
End Sub
