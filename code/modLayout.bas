Option Explicit

' ============================================================
' Text layout engine
'
' Purpose:
'   - Scan all parsed shape rows
'   - Determine one common font size for the whole page
'   - Word-wrap each caption for its own box size
'   - Store calculated layout data back onto each shape dictionary
'
' Notes:
'   - This uses an approximate character-width model.
'   - It is intentionally deterministic and simple.
'   - SVG itself does not provide reliable automatic text wrapping,
'     so the SVG builder must output <tspan> lines.
' ============================================================

Public Sub AnalyseTextLayout(ByVal shapes As Collection)

    Dim globalFontSize As Long
    Dim i As Long
    Dim shape As Object
    Dim wrappedLines As Collection

    globalFontSize = DetermineGlobalFontSize(shapes)

    For i = 1 To shapes.Count

        Set shape = shapes(i)

        Set wrappedLines = WrapCaptionToLines( _
            CStr(shape(KEY_CAPTION)), _
            CDbl(shape(KEY_SIZEX)), _
            globalFontSize _
        )

        shape(KEY_FONT_SIZE) = globalFontSize

        If shape.Exists(KEY_WRAPPED_LINES) Then
            shape.Remove KEY_WRAPPED_LINES
        End If

        shape.Add KEY_WRAPPED_LINES, wrappedLines

    Next i

End Sub


Private Function DetermineGlobalFontSize(ByVal shapes As Collection) As Long

    Dim candidateFontSize As Long

    For candidateFontSize = MAX_FONT_SIZE To MIN_FONT_SIZE Step -1

        If AllShapesFitAtFontSize(shapes, candidateFontSize) Then
            DetermineGlobalFontSize = candidateFontSize
            Exit Function
        End If

    Next candidateFontSize

    DetermineGlobalFontSize = MIN_FONT_SIZE

End Function


Private Function AllShapesFitAtFontSize( _
    ByVal shapes As Collection, _
    ByVal fontSize As Long) As Boolean

    Dim i As Long
    Dim shape As Object

    For i = 1 To shapes.Count

        Set shape = shapes(i)

        If Not ShapeCaptionFitsAtFontSize(shape, fontSize) Then
            AllShapesFitAtFontSize = False
            Exit Function
        End If

    Next i

    AllShapesFitAtFontSize = True

End Function


Private Function ShapeCaptionFitsAtFontSize( _
    ByVal shape As Object, _
    ByVal fontSize As Long) As Boolean

    Dim wrappedLines As Collection
    Dim availableHeight As Double
    Dim requiredHeight As Double
    Dim lineHeight As Double

    Set wrappedLines = WrapCaptionToLines( _
        CStr(shape(KEY_CAPTION)), _
        CDbl(shape(KEY_SIZEX)), _
        fontSize _
    )

    lineHeight = CDbl(fontSize) * LINE_HEIGHT_FACTOR

    availableHeight = CDbl(shape(KEY_SIZEY)) - (2 * TEXT_MARGIN_Y)
    requiredHeight = wrappedLines.Count * lineHeight

    ShapeCaptionFitsAtFontSize = (requiredHeight <= availableHeight)

End Function


Public Function WrapCaptionToLines( _
    ByVal caption As String, _
    ByVal boxWidth As Double, _
    ByVal fontSize As Long) As Collection

    Dim result As New Collection
    Dim words As Variant
    Dim i As Long
    Dim wordText As String
    Dim currentLine As String
    Dim testLine As String
    Dim maxCharsPerLine As Long

    caption = Trim$(caption)

    If Len(caption) = 0 Then
        result.Add vbNullString
        Set WrapCaptionToLines = result
        Exit Function
    End If

    maxCharsPerLine = EstimateMaxCharsPerLine(boxWidth, fontSize)

    words = Split(caption, " ")
    currentLine = vbNullString

    For i = LBound(words) To UBound(words)

        wordText = Trim$(CStr(words(i)))

        If Len(wordText) = 0 Then
            GoTo ContinueWordLoop
        End If

        If Len(wordText) > maxCharsPerLine Then

            If Len(currentLine) > 0 Then
                result.Add currentLine
                currentLine = vbNullString
            End If

            AddLongWordChunks result, wordText, maxCharsPerLine

        Else

            If Len(currentLine) = 0 Then
                testLine = wordText
            Else
                testLine = currentLine & " " & wordText
            End If

            If Len(testLine) <= maxCharsPerLine Then
                currentLine = testLine
            Else
                result.Add currentLine
                currentLine = wordText
            End If

        End If

ContinueWordLoop:
    Next i

    If Len(currentLine) > 0 Then
        result.Add currentLine
    End If

    If result.Count = 0 Then
        result.Add vbNullString
    End If

    Set WrapCaptionToLines = result

End Function


Private Function EstimateMaxCharsPerLine( _
    ByVal boxWidth As Double, _
    ByVal fontSize As Long) As Long

    Dim availableWidth As Double
    Dim approxCharWidth As Double
    Dim result As Long

    availableWidth = boxWidth - (2 * TEXT_MARGIN_X)

    If availableWidth < 1 Then
        availableWidth = 1
    End If

    approxCharWidth = CDbl(fontSize) * APPROX_CHAR_WIDTH_FACTOR

    If approxCharWidth < 1 Then
        approxCharWidth = 1
    End If

    result = CLng(Int(availableWidth / approxCharWidth))

    If result < 1 Then
        result = 1
    End If

    EstimateMaxCharsPerLine = result

End Function


Private Sub AddLongWordChunks( _
    ByVal result As Collection, _
    ByVal wordText As String, _
    ByVal maxCharsPerLine As Long)

    Dim remainingText As String
    Dim chunkText As String

    remainingText = wordText

    Do While Len(remainingText) > 0

        chunkText = Left$(remainingText, maxCharsPerLine)
        result.Add chunkText

        If Len(remainingText) <= maxCharsPerLine Then
            remainingText = vbNullString
        Else
            remainingText = Mid$(remainingText, maxCharsPerLine + 1)
        End If

    Loop

End Sub
