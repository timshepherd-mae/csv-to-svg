Public Function BuildSvgText( _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object) As String

    Dim sb As String
    Dim i As Long
    Dim shape As Object

    Dim canvasWidth As Long
    Dim canvasHeight As Long
    Dim pageTitle As String
    Dim backgroundColour As String
    Dim titleColour As String

    canvasWidth = CLng(pageSettings(KEY_PAGEWIDTH))
    canvasHeight = CLng(pageSettings(KEY_PAGEHEIGHT))

    pageTitle = CStr(pageSettings(KEY_PAGETITLE))
    backgroundColour = CStr(pageSettings(KEY_BACKGROUND))
    titleColour = CStr(pageSettings(KEY_TITLECOLOUR))

    sb = sb & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf

    sb = sb & "<svg xmlns=""http://www.w3.org/2000/svg""" & vbCrLf
    sb = sb & "     xmlns:xlink=""http://www.w3.org/1999/xlink""" & vbCrLf

    sb = sb & "     width=""100%""" & vbCrLf
    sb = sb & "     height=""100%""" & vbCrLf
    sb = sb & "     viewBox=""0 0 " & canvasWidth & " " & canvasHeight & """" & vbCrLf
    sb = sb & "     preserveAspectRatio=""xMinYMin meet"">" & vbCrLf

    sb = sb & "  <title>" & XmlSafeText(pageTitle) & "</title>" & vbCrLf

    sb = sb & "  <style>" & vbCrLf
    sb = sb & "    text { font-family: Arial, Helvetica, sans-serif; font-size: 18px; dominant-baseline: middle; text-anchor: middle; }" & vbCrLf
    sb = sb & "    .page-title { font-size: 34px; font-weight: bold; text-anchor: start; dominant-baseline: hanging; }" & vbCrLf
    sb = sb & "    rect { stroke-width: 2; }" & vbCrLf
    sb = sb & "  </style>" & vbCrLf

    sb = sb & "  <rect x=""0"" y=""0""" & _
              " width=""" & canvasWidth & """" & _
              " height=""" & canvasHeight & """" & _
              " fill=""" & XmlSafeAttribute(backgroundColour) & """ />" & vbCrLf

    If Len(Trim$(pageTitle)) > 0 Then
        sb = sb & "  <text class=""page-title""" & _
                  " x=""40""" & _
                  " y=""30""" & _
                  " fill=""" & XmlSafeAttribute(titleColour) & """>" & _
                  XmlSafeText(pageTitle) & _
                  "</text>" & vbCrLf
    End If

    For i = 1 To shapes.Count
        Set shape = shapes(i)
        sb = sb & BuildShapeSvg(shape)
    Next i

    sb = sb & "</svg>" & vbCrLf

    BuildSvgText = sb

End Function

Public Sub GetSvgExtents(ByVal shapes As Collection, _
                          ByRef minX As Double, _
                          ByRef minY As Double, _
                          ByRef width As Double, _
                          ByRef height As Double)

    Dim i As Long
    Dim shape As Object
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim maxX As Double
    Dim maxY As Double

    Set shape = shapes(1)

    minX = CDbl(shape(KEY_POSX))
    minY = CDbl(shape(KEY_POSY))
    maxX = CDbl(shape(KEY_POSX)) + CDbl(shape(KEY_SIZEX))
    maxY = CDbl(shape(KEY_POSY)) + CDbl(shape(KEY_SIZEY))

    For i = 2 To shapes.Count

        Set shape = shapes(i)

        x1 = CDbl(shape(KEY_POSX))
        y1 = CDbl(shape(KEY_POSY))
        x2 = x1 + CDbl(shape(KEY_SIZEX))
        y2 = y1 + CDbl(shape(KEY_SIZEY))

        If x1 < minX Then minX = x1
        If y1 < minY Then minY = y1
        If x2 > maxX Then maxX = x2
        If y2 > maxY Then maxY = y2

    Next i

    width = maxX - minX
    height = maxY - minY

End Sub



