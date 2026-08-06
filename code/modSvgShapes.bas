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

    sb = sb & "      <text x=""" & SvgNum(cx) & _
              """ y=""" & SvgNum(cy) & _
              """ fill=""" & XmlSafeAttribute(CStr(shape(KEY_TEXTCOL))) & _
              """>" & _
              XmlSafeText(CStr(shape(KEY_CAPTION))) & _
              "</text>" & vbCrLf

    sb = sb & "    </g>" & vbCrLf
    sb = sb & "  </a>" & vbCrLf

    BuildShapeSvg = sb

End Function



