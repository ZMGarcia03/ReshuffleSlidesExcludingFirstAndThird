Sub ReshuffleSlidesExcludingFirstAndThird()
    Dim ppt As Object
    Dim slidesCount As Integer
    Dim slideIndices() As Integer
    Dim i As Integer, j As Integer
    Dim tempIndex As Integer
    
    ' Set PowerPoint Application
    Set ppt = CreateObject("PowerPoint.Application")
    
    ' Change the path to your presentation file
    ppt.Presentations.Open "C:\Path\To\Your\Presentation.pptx"
    
    ' Get the total number of slides
    slidesCount = ppt.ActivePresentation.Slides.Count
    
    ' Initialize the array of slide indices excluding the 1st and 3rd slides
    ReDim slideIndices(2 To slidesCount)
    
    j = 2
    For i = 1 To slidesCount
        If i <> 1 And i <> 3 Then
            slideIndices(j) = i
            j = j + 1
        End If
    Next i
    
    ' Shuffle the slide indices
    For i = UBound(slideIndices) To 2 Step -1
        j = Int((i - 1) * Rnd + 2)
        tempIndex = slideIndices(i)
        slideIndices(i) = slideIndices(j)
        slideIndices(j) = tempIndex
    Next i
    
    ' Apply the new slide order
    For i = 2 To slidesCount
        ppt.ActivePresentation.Slides(i).MoveTo slideIndices(i)
    Next i
    
    ' Save and close the presentation
    ppt.ActivePresentation.Save
    ppt.Quit
    
    ' Release the PowerPoint Application object
    Set ppt = Nothing
End Sub
