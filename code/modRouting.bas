Option Explicit

' ============================================================
' Routing Proof-of-Concept Parser
'
' Purpose:
'   - Convert a compact route code into a list of route points.
'   - Use shape socket lanes and corridor lanes.
'   - Do not render SVG yet.
'
' Route example:
'
'   ID3-W3|L1-3|D3-1|R1-3|ID9-N3
'
' Meaning:
'
'   ID3-W3
'       Start at shape ID3, socket lane W3
'
'   L1-3
'       Move left 1 grid step, stopping on corridor lane 3
'
'   D3-1
'       Move down 3 grid steps, stopping on corridor lane 1
'
'   R1-3
'       Move right 1 grid step, stopping on corridor lane 3
'
'   ID9-N3
'       Terminate at shape ID9, socket lane N3
'
' Notes:
'   - This parser assumes shapes already have derived KEY_POSX and KEY_POSY.
'   - This parser assumes shapes also retain KEY_GRIDCOL and KEY_GRIDROW.
'   - hGrid and vGrid should contain the shape-axis coordinates.
'   - hGrid/vGrid can be produced from strings such as:
'       "50|520|990|1460"
'       "100|260|420|580|740"
' ============================================================


' ============================================================
' Public entry point
' ============================================================

Public Function ParseRouteCode( _
    ByVal routeCode As String, _
    ByVal shapes As Collection, _
    ByVal hGrid As Variant, _
    ByVal vGrid As Variant) As Collection

    Dim routePoints As New Collection
    Dim tokens As Variant
    Dim tokenIndex As Long

    Dim startInfo As Object
    Dim endInfo As Object
    Dim moveInfo As Object

    Dim currentX As Double
    Dim currentY As Double
    Dim currentGridCol As Long
    Dim currentGridRow As Long

    Dim nextX As Double
    Dim nextY As Double

    Dim tokenText As String

    routeCode = Trim$(routeCode)

    If Len(routeCode) = 0 Then
        Err.Raise vbObjectError + 3000, "ParseRouteCode", "Route code is blank."
    End If

    tokens = Split(routeCode, "|")

    If UBound(tokens) < 1 Then
        Err.Raise vbObjectError + 3001, "ParseRouteCode", _
                  "Route code must contain at least a start and end token."
    End If

    ' --------------------------------------------------------
    ' Start token, e.g. ID3-W3
    ' --------------------------------------------------------

    Set startInfo = ResolveEndpointToken(CStr(tokens(0)), shapes)

    currentX = CDbl(startInfo("x"))
    currentY = CDbl(startInfo("y"))
    currentGridCol = CLng(startInfo("gridcol"))
    currentGridRow = CLng(startInfo("gridrow"))

    routePoints.Add CreateRoutePoint( _
        currentX, _
        currentY, _
        "START " & CStr(tokens(0)) _
    )

    ' --------------------------------------------------------
    ' Middle movement tokens, e.g. L1-3, D3-1, R1-3
    ' --------------------------------------------------------

    For tokenIndex = 1 To UBound(tokens) - 1

        tokenText = Trim$(CStr(tokens(tokenIndex)))

        If Len(tokenText) = 0 Then
            GoTo ContinueTokenLoop
        End If

        Set moveInfo = ParseMoveToken(tokenText)

        Select Case UCase$(CStr(moveInfo("direction")))

            Case "L", "R"

                nextX = ResolveVerticalCorridorLaneX( _
                    hGrid, _
                    currentGridCol, _
                    CLng(moveInfo("steps")), _
                    CStr(moveInfo("direction")), _
                    CLng(moveInfo("corridorlane")) _
                )

                nextY = currentY

                currentGridCol = ApplyHorizontalGridStep( _
                    currentGridCol, _
                    CLng(moveInfo("steps")), _
                    CStr(moveInfo("direction")) _
                )

            Case "U", "D"

                nextX = currentX

                nextY = ResolveHorizontalCorridorLaneY( _
                    vGrid, _
                    currentGridRow, _
                    CLng(moveInfo("steps")), _
                    CStr(moveInfo("direction")), _
                    CLng(moveInfo("corridorlane")) _
                )

                currentGridRow = ApplyVerticalGridStep( _
                    currentGridRow, _
                    CLng(moveInfo("steps")), _
                    CStr(moveInfo("direction")) _
                )

            Case Else

                Err.Raise vbObjectError + 3002, "ParseRouteCode", _
                          "Unsupported route direction: " & CStr(moveInfo("direction"))

        End Select

        routePoints.Add CreateRoutePoint( _
            nextX, _
            nextY, _
            "MOVE " & tokenText _
        )

        currentX = nextX
        currentY = nextY

ContinueTokenLoop:
    Next tokenIndex

    ' --------------------------------------------------------
    ' End token, e.g. ID9-N3
    ' --------------------------------------------------------

    Set endInfo = ResolveEndpointToken(CStr(tokens(UBound(tokens))), shapes)

    routePoints.Add CreateRoutePoint( _
        CDbl(endInfo("x")), _
        CDbl(endInfo("y")), _
        "END " & CStr(tokens(UBound(tokens))) _
    )

    Set ParseRouteCode = routePoints

End Function


' ============================================================
' Optional helper for testing in Immediate Window
' ============================================================

Public Sub DebugRouteCode( _
    ByVal routeCode As String, _
    ByVal shapes As Collection, _
    ByVal hGrid As Variant, _
    ByVal vGrid As Variant)

    Dim points As Collection
    Dim p As Object
    Dim i As Long

    Set points = ParseRouteCode(routeCode, shapes, hGrid, vGrid)

    Debug.Print "ROUTE: " & routeCode
    Debug.Print "POINT COUNT: " & points.Count

    For i = 1 To points.Count

        Set p = points(i)

        Debug.Print _
            CStr(i) & ": " & _
            CStr(p("description")) & _
            "  x=" & SvgNum(CDbl(p("x"))) & _
            "  y=" & SvgNum(CDbl(p("y")))

    Next i

End Sub


' ============================================================
' Axis parsing helper
'
' This is useful if your _META values are stored as:
'   hgrid = 50|520|990|1460
'   vgrid = 100|260|420|580|740
'
' It returns a zero-based array of Doubles.
' Grid index 1 maps to array index 0.
' ============================================================

Public Function ParseGridAxisList(ByVal axisText As String) As Variant

    Dim parts As Variant
    Dim result() As Double
    Dim i As Long

    axisText = Trim$(axisText)

    If Len(axisText) = 0 Then
        Err.Raise vbObjectError + 3010, "ParseGridAxisList", "Grid axis list is blank."
    End If

    parts = Split(axisText, "|")

    ReDim result(0 To UBound(parts))

    For i = LBound(parts) To UBound(parts)
        result(i) = ToDbl(CStr(parts(i)))
    Next i

    ParseGridAxisList = result

End Function


' ============================================================
' Endpoint parsing
'
' Endpoint example:
'   ID3-W3
'
' Shape:
'   ID3
'
' Socket lane:
'   W3
' ============================================================

Private Function ResolveEndpointToken( _
    ByVal endpointToken As String, _
    ByVal shapes As Collection) As Object

    Dim result As Object
    Dim parts As Variant
    Dim shapeId As String
    Dim socketId As String
    Dim shape As Object
    Dim socketPoint As Object

    endpointToken = Trim$(endpointToken)

    parts = Split(endpointToken, "-")

    If UBound(parts) <> 1 Then
        Err.Raise vbObjectError + 3020, "ResolveEndpointToken", _
                  "Endpoint token must be in the form ID3-W3: " & endpointToken
    End If

    shapeId = Trim$(CStr(parts(0)))
    socketId = Trim$(CStr(parts(1)))

    Set shape = FindShapeByRouteId(shapes, shapeId)
    Set socketPoint = ResolveSocketLane(shape, socketId)

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("shapeid") = shapeId
    result("socketid") = socketId
    result("x") = socketPoint("x")
    result("y") = socketPoint("y")
    result("gridcol") = CLng(shape(KEY_GRIDCOL))
    result("gridrow") = CLng(shape(KEY_GRIDROW))

    Set ResolveEndpointToken = result

End Function


Private Function FindShapeByRouteId( _
    ByVal shapes As Collection, _
    ByVal routeShapeId As String) As Object

    Dim shape As Object
    Dim wantedId As String
    Dim actualId As String

    wantedId = NormaliseRouteShapeId(routeShapeId)

    For Each shape In shapes

        actualId = NormaliseRouteShapeId(CStr(shape(KEY_ID)))

        If StrComp(actualId, wantedId, vbTextCompare) = 0 Then
            Set FindShapeByRouteId = shape
            Exit Function
        End If

    Next shape

    Err.Raise vbObjectError + 3021, "FindShapeByRouteId", _
              "Shape not found for route ID: " & routeShapeId

End Function


Private Function NormaliseRouteShapeId(ByVal shapeId As String) As String

    shapeId = Trim$(shapeId)

    If UCase$(Left$(shapeId, 2)) = "ID" Then
        shapeId = Mid$(shapeId, 3)
    End If

    NormaliseRouteShapeId = shapeId

End Function


' ============================================================
' Socket lane resolution
'
' Socket lane IDs:
'   N1 N2 N3
'   E1 E2 E3
'   S1 S2 S3
'   W1 W2 W3
'
' For rectangular shapes:
'
'   N1/N2/N3 and S1/S2/S3 divide the width into quarters.
'   W1/W2/W3 and E1/E2/E3 divide the height into quarters.
'
' Example:
'   N1 = 25% along top edge
'   N2 = 50% along top edge
'   N3 = 75% along top edge
' ============================================================

Private Function ResolveSocketLane( _
    ByVal shape As Object, _
    ByVal socketId As String) As Object

    Dim sideCode As String
    Dim socketLane As Long

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    Dim px As Double
    Dim py As Double

    socketId = Trim$(socketId)

    If Len(socketId) < 2 Then
        Err.Raise vbObjectError + 3030, "ResolveSocketLane", _
                  "Invalid socket lane ID: " & socketId
    End If

    sideCode = UCase$(Left$(socketId, 1))
    socketLane = CLng(Mid$(socketId, 2))

    ValidateLaneNumber socketLane, "socket lane", socketId

    x = CDbl(shape(KEY_POSX))
    y = CDbl(shape(KEY_POSY))
    w = CDbl(shape(KEY_SIZEX))
    h = CDbl(shape(KEY_SIZEY))

    Select Case sideCode

        Case "N"
            px = x + ((w / 4) * socketLane)
            py = y

        Case "S"
            px = x + ((w / 4) * socketLane)
            py = y + h

        Case "W"
            px = x
            py = y + ((h / 4) * socketLane)

        Case "E"
            px = x + w
            py = y + ((h / 4) * socketLane)

        Case Else
            Err.Raise vbObjectError + 3031, "ResolveSocketLane", _
                      "Invalid socket side: " & sideCode

    End Select

    Set ResolveSocketLane = CreateRoutePoint( _
        px, _
        py, _
        "SOCKET " & socketId _
    )

End Function


' ============================================================
' Move token parsing
'
' Move token example:
'   L1-3
'
' Meaning:
'   Direction      = L
'   Steps          = 1
'   Corridor lane  = 3
'
' Valid directions:
'   L, R, U, D
' ============================================================

Private Function ParseMoveToken(ByVal moveToken As String) As Object

    Dim result As Object
    Dim directionCode As String
    Dim bodyText As String
    Dim parts As Variant
    Dim steps As Long
    Dim corridorLane As Long

    moveToken = Trim$(moveToken)

    If Len(moveToken) < 4 Then
        Err.Raise vbObjectError + 3040, "ParseMoveToken", _
                  "Invalid move token: " & moveToken
    End If

    directionCode = UCase$(Left$(moveToken, 1))

    If InStr(1, "LRUD", directionCode, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 3041, "ParseMoveToken", _
                  "Invalid move direction: " & directionCode
    End If

    bodyText = Mid$(moveToken, 2)
    parts = Split(bodyText, "-")

    If UBound(parts) <> 1 Then
        Err.Raise vbObjectError + 3042, "ParseMoveToken", _
                  "Move token must be Direction + Steps + '-' + CorridorLane, e.g. L1-3: " & moveToken
    End If

    steps = CLng(parts(0))
    corridorLane = CLng(parts(1))

    If steps < 1 Then
        Err.Raise vbObjectError + 3043, "ParseMoveToken", _
                  "Move steps must be 1 or greater: " & moveToken
    End If

    ValidateLaneNumber corridorLane, "corridor lane", moveToken

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("direction") = directionCode
    result("steps") = steps
    result("corridorlane") = corridorLane

    Set ParseMoveToken = result

End Function


' ============================================================
' Corridor lane resolution
'
' Vertical corridor lanes:
'   Used when a move is L or R.
'   The corridor lies between two hGrid shape axes.
'
' Horizontal corridor lanes:
'   Used when a move is U or D.
'   The corridor lies between two vGrid shape axes.
'
' Lane numbering is always absolute within the gap:
'
'   For vertical corridor lanes:
'       lane 1 = left side of gap
'       lane 2 = centre of gap
'       lane 3 = right side of gap
'
'   For horizontal corridor lanes:
'       lane 1 = top side of gap
'       lane 2 = centre of gap
'       lane 3 = lower side of gap
' ============================================================

Private Function ResolveVerticalCorridorLaneX( _
    ByVal hGrid As Variant, _
    ByVal currentGridCol As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String, _
    ByVal corridorLane As Long) As Double

    Dim targetGridCol As Long
    Dim leftCol As Long
    Dim rightCol As Long

    directionCode = UCase$(directionCode)

    If directionCode = "L" Then
        targetGridCol = currentGridCol - steps
    ElseIf directionCode = "R" Then
        targetGridCol = currentGridCol + steps
    Else
        Err.Raise vbObjectError + 3050, "ResolveVerticalCorridorLaneX", _
                  "Direction must be L or R."
    End If

    If targetGridCol < 1 Then
        Err.Raise vbObjectError + 3051, "ResolveVerticalCorridorLaneX", _
                  "Horizontal move goes before first grid column."
    End If

    leftCol = MinLong(currentGridCol, targetGridCol)
    rightCol = MaxLong(currentGridCol, targetGridCol)

    ResolveVerticalCorridorLaneX = InterpolateLane( _
        GetGridAxisValue(hGrid, leftCol), _
        GetGridAxisValue(hGrid, rightCol), _
        corridorLane _
    )

End Function


Private Function ResolveHorizontalCorridorLaneY( _
    ByVal vGrid As Variant, _
    ByVal currentGridRow As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String, _
    ByVal corridorLane As Long) As Double

    Dim targetGridRow As Long
    Dim topRow As Long
    Dim bottomRow As Long

    directionCode = UCase$(directionCode)

    If directionCode = "U" Then
        targetGridRow = currentGridRow - steps
    ElseIf directionCode = "D" Then
        targetGridRow = currentGridRow + steps
    Else
        Err.Raise vbObjectError + 3060, "ResolveHorizontalCorridorLaneY", _
                  "Direction must be U or D."
    End If

    If targetGridRow < 1 Then
        Err.Raise vbObjectError + 3061, "ResolveHorizontalCorridorLaneY", _
                  "Vertical move goes before first grid row."
    End If

    topRow = MinLong(currentGridRow, targetGridRow)
    bottomRow = MaxLong(currentGridRow, targetGridRow)

    ResolveHorizontalCorridorLaneY = InterpolateLane( _
        GetGridAxisValue(vGrid, topRow), _
        GetGridAxisValue(vGrid, bottomRow), _
        corridorLane _
    )

End Function


Private Function InterpolateLane( _
    ByVal axisA As Double, _
    ByVal axisB As Double, _
    ByVal laneNumber As Long) As Double

    ' Lane 1 = 25%
    ' Lane 2 = 50%
    ' Lane 3 = 75%

    Select Case laneNumber

        Case 1
            InterpolateLane = axisA + ((axisB - axisA) * 0.25)

        Case 2
            InterpolateLane = axisA + ((axisB - axisA) * 0.5)

        Case 3
            InterpolateLane = axisA + ((axisB - axisA) * 0.75)

        Case Else
            Err.Raise vbObjectError + 3070, "InterpolateLane", _
                      "Lane number must be 1, 2, or 3."

    End Select

End Function


' ============================================================
' Grid state movement
' ============================================================

Private Function ApplyHorizontalGridStep( _
    ByVal currentGridCol As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String) As Long

    directionCode = UCase$(directionCode)

    Select Case directionCode

        Case "L"
            ApplyHorizontalGridStep = currentGridCol - steps

        Case "R"
            ApplyHorizontalGridStep = currentGridCol + steps

        Case Else
            Err.Raise vbObjectError + 3080, "ApplyHorizontalGridStep", _
                      "Direction must be L or R."

    End Select

End Function


Private Function ApplyVerticalGridStep( _
    ByVal currentGridRow As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String) As Long

    directionCode = UCase$(directionCode)

    Select Case directionCode

        Case "U"
            ApplyVerticalGridStep = currentGridRow - steps

        Case "D"
            ApplyVerticalGridStep = currentGridRow + steps

        Case Else
            Err.Raise vbObjectError + 3090, "ApplyVerticalGridStep", _
                      "Direction must be U or D."

    End Select

End Function


' ============================================================
' Axis helpers
' ============================================================

Private Function GetGridAxisValue( _
    ByVal axisValues As Variant, _
    ByVal gridIndex As Long) As Double

    ' Grid index is 1-based.
    ' Array index is assumed to be zero-based.

    Dim arrayIndex As Long

    If gridIndex < 1 Then
        Err.Raise vbObjectError + 3100, "GetGridAxisValue", _
                  "Grid index must be 1 or greater."
    End If

    arrayIndex = gridIndex - 1

    If IsArray(axisValues) Then

        If arrayIndex < LBound(axisValues) Or arrayIndex > UBound(axisValues) Then
            Err.Raise vbObjectError + 3101, "GetGridAxisValue", _
                      "Grid index is outside axis list: " & CStr(gridIndex)
        End If

        GetGridAxisValue = CDbl(axisValues(arrayIndex))
        Exit Function

    End If

    Err.Raise vbObjectError + 3102, "GetGridAxisValue", _
              "Axis values must be supplied as an array."

End Function


' ============================================================
' Point helper
' ============================================================

Private Function CreateRoutePoint( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal description As String) As Object

    Dim point As Object

    Set point = CreateObject("Scripting.Dictionary")
    point.CompareMode = vbTextCompare

    point("x") = x
    point("y") = y
    point("description") = description

    Set CreateRoutePoint = point

End Function


' ============================================================
' Validation/general helpers
' ============================================================

Private Sub ValidateLaneNumber( _
    ByVal laneNumber As Long, _
    ByVal laneType As String, _
    ByVal sourceText As String)

    If laneNumber < 1 Or laneNumber > 3 Then
        Err.Raise vbObjectError + 3110, "ValidateLaneNumber", _
                  laneType & " must be 1, 2, or 3. Source: " & sourceText
    End If

End Sub


Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long

    If a < b Then
        MinLong = a
    Else
        MinLong = b
    End If

End Function


Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long

    If a > b Then
        MaxLong = a
    Else
        MaxLong = b
    End If

End Function


' ============================================================
'
' TEST HARNESS
'
' ============================================================

' ============================================================
' Proof-of-concept runner
'
' Run this procedure directly from the VBA editor.
'
' It creates a small fake shape collection and calls:
'   DebugRouteCode
'
' Output appears in the Immediate Window.
' ============================================================

Public Sub TestRoutingProofOfConcept()

    Dim shapes As Collection
    Dim hGrid As Variant
    Dim vGrid As Variant
    Dim routeCode As String

    Set shapes = New Collection

    hGrid = ParseGridAxisList("50|520|990|1460")
    vGrid = ParseGridAxisList("70|150|275|400|525|650|775")

    ' Test shapes.
    ' These are deliberately minimal dictionaries containing only
    ' the fields required by the routing proof-of-concept.
    shapes.Add CreateTestRoutingShape("3", 3, 1, hGrid, vGrid, 300, 80)
    shapes.Add CreateTestRoutingShape("9", 3, 4, hGrid, vGrid, 300, 80)

    routeCode = "ID3-W3|L1-3|D3-1|R1-3|ID9-N3"

    DebugRouteCode routeCode, shapes, hGrid, vGrid

End Sub

Private Function CreateTestRoutingShape( _
    ByVal idText As String, _
    ByVal gridCol As Long, _
    ByVal gridRow As Long, _
    ByVal hGrid As Variant, _
    ByVal vGrid As Variant, _
    ByVal sizeX As Double, _
    ByVal sizeY As Double) As Object

    Dim shape As Object

    Set shape = CreateObject("Scripting.Dictionary")
    shape.CompareMode = vbTextCompare

    shape(KEY_ID) = idText
    shape(KEY_GRIDCOL) = gridCol
    shape(KEY_GRIDROW) = gridRow

    ' In the current Phase 1 model, hGrid/vGrid values are layout coordinates.
    ' They are used here as the calculated top-left position of the shape.
    shape(KEY_POSX) = CDbl(hGrid(gridCol - 1))
    shape(KEY_POSY) = CDbl(vGrid(gridRow - 1))

    shape(KEY_SIZEX) = sizeX
    shape(KEY_SIZEY) = sizeY

    Set CreateTestRoutingShape = shape

End Function