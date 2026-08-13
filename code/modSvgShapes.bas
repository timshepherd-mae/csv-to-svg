Public Function BuildShapeSvg(ByVal shape As Object) As String

    Dim sb As String
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim r As Double
    Dim cx As Double
    Dim cy As Double
    Dim linkUrl As String
    Dim groupId As String
    Dim targetAttr As String

    x = CDbl(shape(KEY_POSX))
    y = CDbl(shape(KEY_POSY))
    w = CDbl(shape(KEY_SIZEX))
    h = CDbl(shape(KEY_SIZEY))
    r = CDbl(shape(KEY_RAD))

    cx = x + (w / 2)
    cy = y + (h / 2)

    groupId = "shape-" & XmlSafeAttribute(CStr(shape(KEY_ID)))
    linkUrl = CStr(shape(KEY_LINKURL))
    

    If CStr(shape(KEY_NEWPAGE)) = "1" Then
        targetAttr = " target=""_blank"""
    Else
        targetAttr = vbNullString
    End If


    sb = sb & "  <a xlink:href=""" & _
              XmlSafeAttribute(linkUrl) & _
              """" & _
              targetAttr & _
              ">" & vbCrLf

    sb = sb & "    <g id=""" & groupId & """>" & vbCrLf

    sb = sb & "      <title>" & _
              XmlSafeText(CStr(shape(KEY_TOOLTIP))) & _
              "</title>" & vbCrLf

    sb = sb & "      <rect x=""" & SvgNum(x) & _
              """ y=""" & SvgNum(y) & _
              """ width=""" & SvgNum(w) & _
              """ height=""" & SvgNum(h) & _
              """ rx=""" & SvgNum(r) & _
              """ ry=""" & SvgNum(r) & _
              """ fill=""" & XmlSafeAttribute(CStr(shape(KEY_FILLCOL))) & _
              """ stroke=""" & XmlSafeAttribute(CStr(shape(KEY_STROKECOL))) & _
              """ />" & vbCrLf

    sb = sb & BuildWrappedCaptionSvg(shape, cx, cy)

    sb = sb & "    </g>" & vbCrLf
    sb = sb & " </a>" & vbCrLf

    BuildShapeSvg = sb

End Function

Private Function BuildWrappedCaptionSvg( _
    ByVal shape As Object, _
    ByVal cx As Double, _
    ByVal cy As Double) As String

    Dim sb As String
    Dim wrappedLines As Collection
    Dim fontSize As Long
    Dim lineHeight As Double
    Dim firstY As Double
    Dim lineY As Double
    Dim i As Long

    fontSize = CLng(shape(KEY_FONT_SIZE))
    Set wrappedLines = shape(KEY_WRAPPED_LINES)

    lineHeight = CDbl(fontSize) * LINE_HEIGHT_FACTOR

    firstY = cy - (((wrappedLines.Count - 1) * lineHeight) / 2)

    sb = sb & "      <text" & _
              " x=""" & SvgNum(cx) & """" & _
              " y=""" & SvgNum(cy) & """" & _
              " fill=""" & XmlSafeAttribute(CStr(shape(KEY_TEXTCOL))) & """" & _
              " font-size=""" & CStr(fontSize) & """" & _
              " text-anchor=""middle""" & _
              " dominant-baseline=""middle""" & _
              ">" & vbCrLf

    For i = 1 To wrappedLines.Count

        lineY = firstY + ((i - 1) * lineHeight)

        sb = sb & "        <tspan" & _
                  " x=""" & SvgNum(cx) & """" & _
                  " y=""" & SvgNum(lineY) & """" & _
                  ">" & _
                  XmlSafeText(CStr(wrappedLines(i))) & _
                  "</tspan>" & vbCrLf

    Next i

    sb = sb & "      </text>" & vbCrLf

    BuildWrappedCaptionSvg = sb

End Function



