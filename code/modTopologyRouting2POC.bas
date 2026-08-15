Option Explicit

' ============================================================
' modTopologyRouting2POC
'
' Concept 2 topology-routing proof of concept.
'
' Purpose:
'   - Test topology-based route parsing.
'   - Output route vertices with explicit vertex types.
'   - Do not render SVG yet.
'
' Concept 2 route grammar:
'
'   START TOKEN:
'       ID3-W3-3
'
'       Shape ID      = ID3
'       Socket        = W3
'       Route lane    = 3
'
'       If route lane is omitted:
'
'       ID3-W3
'
'       then route lane defaults to 2.
'
'   MOVE TOKEN:
'       D3-1
'
'       Direction     = D
'       Steps         = 3
'       Route lane    = 1
'
'       If route lane is omitted:
'
'       D3
'
'       then route lane defaults to 2.
'
'   END TOKEN:
'       ID9-N3
'
'       Shape ID      = ID9
'       Socket        = N3
'
' Key topology rule:
'
'   The start socket determines the initial exit direction.
'
'       W socket exits left
'       E socket exits right
'       N socket exits up
'       S socket exits down
'
'   The destination socket determines the final entry direction.
'
'       N socket is entered from above
'       S socket is entered from below
'       W socket is entered from the left
'       E socket is entered from the right
'
' Vertex types:
'
'   SOCKET_START
'   ROUTE_ENTRY
'   ROUTE_BEND
'   ROUTE_EXIT
'   SOCKET_END
'
' Test routes from current diagram:
'
'   ROUTE 1:
'       ID3-W3-3|D3-1|R1-3|ID9-N3
'
'   ROUTE 2:
'       ID3-W1-1|D5-2|R1-1|ID9-S1
'
'   ROUTE 3:
'       ID3-S2-3|ID6-N2
' ============================================================

Private Const DEFAULT_ROUTE_LANE As Long = 2

Private Const VERTEX_SOCKET_START As String = "SOCKET_START"
Private Const VERTEX_ROUTE_ENTRY As String = "ROUTE_ENTRY"
Private Const VERTEX_ROUTE_VERTEX As String = "ROUTE_VERTEX"
Private Const VERTEX_ROUTE_EXIT As String = "ROUTE_EXIT"
Private Const VERTEX_SOCKET_END As String = "SOCKET_END"


' ============================================================
' Public test harness
'
' Run this procedure directly.
' Open Immediate Window with Ctrl + G.
' ============================================================

Public Sub TestTopologyRouting2POC()

    Dim hGrid As Variant
    Dim vGrid As Variant
    Dim shapes As Collection

    Set shapes = New Collection

    hGrid = Array(50#, 520#, 990#, 1460#)
    vGrid = Array(70#, 150#, 275#, 400#, 525#, 650#, 775#)

    ' Shape positions use current convention:
    ' x/y represent top-left corner of the shape.

    shapes.Add Topo2_CreateShape("3", 3, 1, hGrid, vGrid, 300, 80)
    shapes.Add Topo2_CreateShape("6", 3, 3, hGrid, vGrid, 300, 80)
    shapes.Add Topo2_CreateShape("9", 3, 4, hGrid, vGrid, 300, 80)

    Debug.Print
    Debug.Print "=================================================="
    Debug.Print "TOPOLOGY ROUTING 2 POC - EXPLICIT VERTEX TYPES"
    Debug.Print "=================================================="

    Topo2_TestRoute _
        "ROUTE 1", _
        "ID3-W3-3|D3-1|R1-3|ID9-N3", _
        shapes, _
        hGrid, _
        vGrid

    Topo2_TestRoute _
        "ROUTE 2", _
        "ID3-W1-1|D5-2|R1-1|ID9-S1", _
        shapes, _
        hGrid, _
        vGrid

    Topo2_TestRoute _
        "ROUTE 3", _
        "ID3-S2-3|ID6-N2", _
        shapes, _
        hGrid, _
        vGrid

End Sub

Public Sub TestTopologyRouting2POC_RealData( _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object)

    Dim hGrid As Variant
    Dim vGrid As Variant

    hGrid = ParseGridAxisList( _
                CStr(pageSettings(KEY_HGRID)))

    vGrid = ParseGridAxisList( _
                CStr(pageSettings(KEY_VGRID)))

    Debug.Print
    Debug.Print
    Debug.Print "========================================"
    Debug.Print "TOPOLOGY ROUTING USING REAL SHAPES"
    Debug.Print "========================================"

    Topo2_TestRoute _
        "ROUTE 1", _
        "Ra1-E3-1|D1-1|Ra2-E1", _
        shapes, _
        hGrid, _
        vGrid

    Topo2_TestRoute _
        "ROUTE 2", _
        "Ra3-E2-2|D3-2|R2-2|U5-2|Rc2-W2", _
        shapes, _
        hGrid, _
        vGrid

End Sub


Private Sub Topo2_TestRoute( _
    ByVal routeName As String, _
    ByVal routeCode As String, _
    ByVal shapes As Collection, _
    ByVal hGrid As Variant, _
    ByVal vGrid As Variant)

    Dim points As Collection
    Dim i As Long
    Dim p As Object

    Set points = Topo2_ParseRouteCode(routeCode, shapes, hGrid, vGrid)

    Debug.Print
    Debug.Print "--------------------------------------------------"
    Debug.Print routeName
    Debug.Print routeCode
    Debug.Print "POINT COUNT: " & points.Count
    Debug.Print "--------------------------------------------------"

    For i = 1 To points.Count

        Set p = points(i)

        Debug.Print _
            CStr(i) & ": " & _
            CStr(p("vertextype")) & _
            "  x=" & Topo2_Num(CDbl(p("x"))) & _
            "  y=" & Topo2_Num(CDbl(p("y")))

        Debug.Print _
            "       " & _
            CStr(p("label"))

    Next i

    Topo2_PrintRelationshipChecks routeName, points

End Sub


' ============================================================
' Main parser
' ============================================================

Public Function Topo2_ParseRouteCode( _
    ByVal routeCode As String, _
    ByVal shapes As Collection, _
    ByVal hGrid As Variant, _
    ByVal vGrid As Variant) As Collection

    Dim routePoints As New Collection
    Dim tokens As Variant
    Dim tokenIndex As Long
    Dim middleMoveCount As Long

    Dim startInfo As Object
    Dim endInfo As Object
    Dim moveInfo As Object

    Dim currentX As Double
    Dim currentY As Double
    Dim currentGridCol As Long
    Dim currentGridRow As Long

    Dim startExitDirection As String
    Dim startRouteLane As Long

    Dim directionCode As String
    Dim stepCount As Long
    Dim routeLane As Long
    Dim isFinalMove As Boolean

    Dim vertexX As Double
    Dim vertexY As Double

    routeCode = Trim$(routeCode)

    If Len(routeCode) = 0 Then
        Err.Raise vbObjectError + 5000, "Topo2_ParseRouteCode", "Route code is blank."
    End If

    tokens = Split(routeCode, "|")

    If UBound(tokens) < 1 Then
        Err.Raise vbObjectError + 5001, "Topo2_ParseRouteCode", _
                  "Route code must include start and end tokens."
    End If

    middleMoveCount = UBound(tokens) - 1

    Set startInfo = Topo2_ParseStartToken(CStr(tokens(0)), shapes)
    Set endInfo = Topo2_ParseEndToken(CStr(tokens(UBound(tokens))), shapes)

    currentX = CDbl(startInfo("x"))
    currentY = CDbl(startInfo("y"))
    currentGridCol = CLng(startInfo("gridcol"))
    currentGridRow = CLng(startInfo("gridrow"))

    routePoints.Add Topo2_CreatePoint( _
        currentX, _
        currentY, _
        VERTEX_SOCKET_START, _
        "START " & CStr(tokens(0)) _
    )

    ' --------------------------------------------------------
    ' Implicit socket exit.
    '
    ' This creates point A.
    '
    ' The socket side decides the exit direction.
    ' The start token's optional final lane value decides the
    ' route corridor lane.
    ' --------------------------------------------------------

    startExitDirection = Topo2_ExitDirectionFromSocket(CStr(startInfo("socketid")))
    startRouteLane = CLng(startInfo("routelane"))

    If middleMoveCount = 0 Then

        ' ----------------------------------------------------
        ' Direct route.
        '
        ' Example:
        '   ID3-S2-3|ID6-N2
        '
        ' There are no explicit middle moves, so A is placed
        ' between the start and end sockets on the relevant axis.
        ' ----------------------------------------------------

        Call Topo2_ResolveDirectRouteEntryPoint( _
            startInfo, _
            endInfo, _
            startExitDirection, _
            startRouteLane, _
            vertexX, _
            vertexY _
        )

    Else

        ' ----------------------------------------------------
        ' Standard route with middle moves.
        ' ----------------------------------------------------

        Select Case startExitDirection

            Case "L", "R"

                vertexX = Topo2_ResolveVerticalRouteCorridorX( _
                    hGrid, _
                    currentGridCol, _
                    1, _
                    startExitDirection, _
                    startRouteLane _
                )

                vertexY = currentY

                currentGridCol = Topo2_ApplyHorizontalGridStep( _
                    currentGridCol, _
                    1, _
                    startExitDirection _
                )

            Case "U", "D"

                vertexX = currentX

                vertexY = Topo2_ResolveHorizontalRouteCorridorY( _
                    vGrid, _
                    currentGridRow, _
                    1, _
                    startExitDirection, _
                    startRouteLane _
                )

                currentGridRow = Topo2_ApplyVerticalGridStep( _
                    currentGridRow, _
                    1, _
                    startExitDirection _
                )

            Case Else

                Err.Raise vbObjectError + 5002, "Topo2_ParseRouteCode", _
                          "Invalid start exit direction."

        End Select

    End If

    routePoints.Add Topo2_CreatePoint( _
        vertexX, _
        vertexY, _
        VERTEX_ROUTE_ENTRY, _
        "A IMPLICIT_EXIT_TO_ROUTE_LANE_" & CStr(startRouteLane) _
    )

    currentX = vertexX
    currentY = vertexY

    ' --------------------------------------------------------
    ' Middle move tokens.
    '
    ' The final middle move is destination-aware and creates
    ' ROUTE_EXIT.
    '
    ' Earlier middle moves create ROUTE_BEND.
    ' --------------------------------------------------------

    If middleMoveCount > 0 Then

        For tokenIndex = 1 To UBound(tokens) - 1

            Set moveInfo = Topo2_ParseMoveToken(CStr(tokens(tokenIndex)))

            directionCode = CStr(moveInfo("direction"))
            stepCount = CLng(moveInfo("steps"))
            routeLane = CLng(moveInfo("routelane"))

            isFinalMove = (tokenIndex = UBound(tokens) - 1)

            If isFinalMove Then

                ' ------------------------------------------------
                ' Destination-aware final move.
                '
                ' The destination socket side determines whether
                ' the route exit point must share X or share Y
                ' with the destination socket.
                ' ------------------------------------------------

                Select Case Topo2_SocketSide(CStr(endInfo("socketid")))

                    Case "N", "S"

                        vertexX = CDbl(endInfo("x"))
                        vertexY = currentY

                    Case "E", "W"

                        vertexX = currentX
                        vertexY = CDbl(endInfo("y"))

                    Case Else

                        Err.Raise vbObjectError + 5003, "Topo2_ParseRouteCode", _
                                  "Invalid destination socket side."

                End Select

                routePoints.Add Topo2_CreatePoint( _
                    vertexX, _
                    vertexY, _
                    VERTEX_ROUTE_EXIT, _
                    "C FINAL_ROUTE_EXIT " & CStr(tokens(tokenIndex)) _
                )

                currentX = vertexX
                currentY = vertexY

            Else

                ' ------------------------------------------------
                ' Non-final route-corridor travel.
                ' ------------------------------------------------

                Select Case directionCode

                    Case "L", "R"

                        vertexX = Topo2_ResolveVerticalRouteCorridorX( _
                            hGrid, _
                            currentGridCol, _
                            stepCount, _
                            directionCode, _
                            routeLane _
                        )

                        vertexY = currentY

                        currentGridCol = Topo2_ApplyHorizontalGridStep( _
                            currentGridCol, _
                            stepCount, _
                            directionCode _
                        )

                    Case "U", "D"

                        vertexX = currentX

                        vertexY = Topo2_ResolveHorizontalRouteCorridorY( _
                            vGrid, _
                            currentGridRow, _
                            stepCount, _
                            directionCode, _
                            routeLane _
                        )

                        currentGridRow = Topo2_ApplyVerticalGridStep( _
                            currentGridRow, _
                            stepCount, _
                            directionCode _
                        )

                    Case Else

                        Err.Raise vbObjectError + 5004, "Topo2_ParseRouteCode", _
                                  "Invalid move direction: " & directionCode

                End Select

                routePoints.Add Topo2_CreatePoint( _
                    vertexX, _
                    vertexY, _
                    VERTEX_ROUTE_VERTEX, _
                    "B ROUTE_TRAVEL " & CStr(tokens(tokenIndex)) _
                )

                currentX = vertexX
                currentY = vertexY

            End If

        Next tokenIndex

    End If

    routePoints.Add Topo2_CreatePoint( _
        CDbl(endInfo("x")), _
        CDbl(endInfo("y")), _
        VERTEX_SOCKET_END, _
        "END " & CStr(tokens(UBound(tokens))) _
    )

    Set Topo2_ParseRouteCode = routePoints

End Function


' ============================================================
' Token parsing
' ============================================================

Private Function Topo2_ParseStartToken( _
    ByVal tokenText As String, _
    ByVal shapes As Collection) As Object

    Dim result As Object
    Dim parts As Variant
    Dim shapeId As String
    Dim socketId As String
    Dim routeLane As Long
    Dim shape As Object
    Dim socketPoint As Object

    parts = Split(Trim$(tokenText), "-")

    If UBound(parts) < 1 Or UBound(parts) > 2 Then
        Err.Raise vbObjectError + 5010, "Topo2_ParseStartToken", _
                  "Start token must be IDx-SOCKET or IDx-SOCKET-LANE: " & tokenText
    End If

    shapeId = Trim$(CStr(parts(0)))
    socketId = Trim$(CStr(parts(1)))

    If UBound(parts) = 2 Then
        routeLane = CLng(parts(2))
    Else
        routeLane = DEFAULT_ROUTE_LANE
    End If

    Topo2_ValidateLane routeLane, "start route lane", tokenText

    Set shape = Topo2_FindShapeByRouteId(shapes, shapeId)
    Set socketPoint = Topo2_ResolveSocket(shape, socketId)

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("shapeid") = shapeId
    result("socketid") = socketId
    result("routelane") = routeLane
    result("x") = socketPoint("x")
    result("y") = socketPoint("y")
    result("gridcol") = CLng(shape(KEY_GRIDCOL))
    result("gridrow") = CLng(shape(KEY_GRIDROW))

    Set Topo2_ParseStartToken = result

End Function


Private Function Topo2_ParseEndToken( _
    ByVal tokenText As String, _
    ByVal shapes As Collection) As Object

    Dim result As Object
    Dim parts As Variant
    Dim shapeId As String
    Dim socketId As String
    Dim shape As Object
    Dim socketPoint As Object

    parts = Split(Trim$(tokenText), "-")

    If UBound(parts) <> 1 Then
        Err.Raise vbObjectError + 5020, "Topo2_ParseEndToken", _
                  "End token must be IDx-SOCKET: " & tokenText
    End If

    shapeId = Trim$(CStr(parts(0)))
    socketId = Trim$(CStr(parts(1)))

    Set shape = Topo2_FindShapeByRouteId(shapes, shapeId)
    Set socketPoint = Topo2_ResolveSocket(shape, socketId)

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("shapeid") = shapeId
    result("socketid") = socketId
    result("x") = socketPoint("x")
    result("y") = socketPoint("y")
    result("gridcol") = CLng(shape(KEY_GRIDCOL))
    result("gridrow") = CLng(shape(KEY_GRIDROW))

    Set Topo2_ParseEndToken = result

End Function


Private Function Topo2_ParseMoveToken(ByVal tokenText As String) As Object

    Dim result As Object
    Dim directionCode As String
    Dim bodyText As String
    Dim parts As Variant
    Dim steps As Long
    Dim routeLane As Long

    tokenText = Trim$(tokenText)

    If Len(tokenText) < 2 Then
        Err.Raise vbObjectError + 5030, "Topo2_ParseMoveToken", _
                  "Invalid move token: " & tokenText
    End If

    directionCode = UCase$(Left$(tokenText, 1))
    bodyText = Mid$(tokenText, 2)

    If InStr(1, "LRUD", directionCode, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5031, "Topo2_ParseMoveToken", _
                  "Invalid move direction: " & directionCode
    End If

    parts = Split(bodyText, "-")

    If UBound(parts) = 0 Then
        steps = CLng(parts(0))
        routeLane = DEFAULT_ROUTE_LANE
    ElseIf UBound(parts) = 1 Then
        steps = CLng(parts(0))
        routeLane = CLng(parts(1))
    Else
        Err.Raise vbObjectError + 5032, "Topo2_ParseMoveToken", _
                  "Move token must be DirectionSteps or DirectionSteps-Lane: " & tokenText
    End If

    If steps < 1 Then
        Err.Raise vbObjectError + 5033, "Topo2_ParseMoveToken", _
                  "Move steps must be 1 or greater: " & tokenText
    End If

    Topo2_ValidateLane routeLane, "move route lane", tokenText

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result("direction") = directionCode
    result("steps") = steps
    result("routelane") = routeLane

    Set Topo2_ParseMoveToken = result

End Function


' ============================================================
' Direct route entry
' ============================================================

Private Sub Topo2_ResolveDirectRouteEntryPoint( _
    ByVal startInfo As Object, _
    ByVal endInfo As Object, _
    ByVal startExitDirection As String, _
    ByVal lane As Long, _
    ByRef vertexX As Double, _
    ByRef vertexY As Double)

    Dim startX As Double
    Dim startY As Double
    Dim endX As Double
    Dim endY As Double

    startX = CDbl(startInfo("x"))
    startY = CDbl(startInfo("y"))
    endX = CDbl(endInfo("x"))
    endY = CDbl(endInfo("y"))

    Select Case startExitDirection

        Case "L", "R"

            vertexX = Topo2_InterpolateLane(startX, endX, lane)
            vertexY = startY

        Case "U", "D"

            vertexX = startX
            vertexY = Topo2_InterpolateLane(startY, endY, lane)

        Case Else

            Err.Raise vbObjectError + 5035, "Topo2_ResolveDirectRouteEntryPoint", _
                      "Invalid start exit direction."

    End Select

End Sub


' ============================================================
' Shape and socket helpers
' ============================================================

Private Function Topo2_CreateShape( _
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
    shape(KEY_POSX) = CDbl(hGrid(gridCol - 1))
    shape(KEY_POSY) = CDbl(vGrid(gridRow - 1))
    shape(KEY_SIZEX) = sizeX
    shape(KEY_SIZEY) = sizeY

    Set Topo2_CreateShape = shape

End Function


Private Function Topo2_FindShapeByRouteId( _
    ByVal shapes As Collection, _
    ByVal routeShapeId As String) As Object

    Dim shape As Object
    Dim targetId As String
    Dim actualId As String

    targetId = Topo2_NormaliseShapeId(routeShapeId)

    For Each shape In shapes

        actualId = Topo2_NormaliseShapeId(CStr(shape(KEY_ID)))

        If StrComp(targetId, actualId, vbTextCompare) = 0 Then
            Set Topo2_FindShapeByRouteId = shape
            Exit Function
        End If

    Next shape

    Err.Raise vbObjectError + 5040, "Topo2_FindShapeByRouteId", _
              "Shape not found: " & routeShapeId

End Function


Private Function Topo2_NormaliseShapeId(ByVal shapeId As String) As String

    shapeId = Trim$(shapeId)

    If UCase$(Left$(shapeId, 2)) = "ID" Then
        shapeId = Mid$(shapeId, 3)
    End If

    Topo2_NormaliseShapeId = shapeId

End Function


Private Function Topo2_ResolveSocket( _
    ByVal shape As Object, _
    ByVal socketId As String) As Object

    Dim sideCode As String
    Dim lane As Long

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    Dim px As Double
    Dim py As Double

    socketId = Trim$(socketId)

    sideCode = UCase$(Left$(socketId, 1))
    lane = CLng(Mid$(socketId, 2))

    Topo2_ValidateLane lane, "socket lane", socketId

    x = CDbl(shape(KEY_POSX))
    y = CDbl(shape(KEY_POSY))
    w = CDbl(shape(KEY_SIZEX))
    h = CDbl(shape(KEY_SIZEY))

    Select Case sideCode

        Case "N"
            px = x + ((w / 4) * lane)
            py = y

        Case "S"
            px = x + ((w / 4) * lane)
            py = y + h

        Case "W"
            px = x
            py = y + ((h / 4) * lane)

        Case "E"
            px = x + w
            py = y + ((h / 4) * lane)

        Case Else
            Err.Raise vbObjectError + 5050, "Topo2_ResolveSocket", _
                      "Invalid socket side: " & socketId

    End Select

    Set Topo2_ResolveSocket = Topo2_CreatePoint( _
        px, _
        py, _
        "SOCKET", _
        "SOCKET " & socketId _
    )

End Function


Private Function Topo2_SocketSide(ByVal socketId As String) As String

    Topo2_SocketSide = UCase$(Left$(Trim$(socketId), 1))

End Function


Private Function Topo2_ExitDirectionFromSocket(ByVal socketId As String) As String

    Select Case Topo2_SocketSide(socketId)

        Case "W"
            Topo2_ExitDirectionFromSocket = "L"

        Case "E"
            Topo2_ExitDirectionFromSocket = "R"

        Case "N"
            Topo2_ExitDirectionFromSocket = "U"

        Case "S"
            Topo2_ExitDirectionFromSocket = "D"

        Case Else
            Err.Raise vbObjectError + 5060, "Topo2_ExitDirectionFromSocket", _
                      "Invalid socket side: " & socketId

    End Select

End Function


' ============================================================
' Corridor coordinate helpers
' ============================================================

Private Function Topo2_ResolveVerticalRouteCorridorX( _
    ByVal hGrid As Variant, _
    ByVal currentGridCol As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String, _
    ByVal lane As Long) As Double

    Dim targetGridCol As Long
    Dim leftCol As Long
    Dim rightCol As Long

    targetGridCol = Topo2_ApplyHorizontalGridStep(currentGridCol, steps, directionCode)

    leftCol = Topo2_MinLong(currentGridCol, targetGridCol)
    rightCol = Topo2_MaxLong(currentGridCol, targetGridCol)

    Topo2_ResolveVerticalRouteCorridorX = Topo2_InterpolateLane( _
        Topo2_GetAxisValue(hGrid, leftCol), _
        Topo2_GetAxisValue(hGrid, rightCol), _
        lane _
    )

End Function


Private Function Topo2_ResolveHorizontalRouteCorridorY( _
    ByVal vGrid As Variant, _
    ByVal currentGridRow As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String, _
    ByVal lane As Long) As Double

    Dim targetGridRow As Long
    Dim topRow As Long
    Dim bottomRow As Long

    targetGridRow = Topo2_ApplyVerticalGridStep(currentGridRow, steps, directionCode)

    If UCase$(directionCode) = "D" Then
        topRow = targetGridRow - 1
        bottomRow = targetGridRow
    Else
        topRow = targetGridRow
        bottomRow = targetGridRow + 1
    End If

    If topRow < 1 Then topRow = 1

    Topo2_ResolveHorizontalRouteCorridorY = Topo2_InterpolateLane( _
        Topo2_GetAxisValue(vGrid, topRow), _
        Topo2_GetAxisValue(vGrid, bottomRow), _
        lane _
    )

End Function


Private Function Topo2_InterpolateLane( _
    ByVal axisA As Double, _
    ByVal axisB As Double, _
    ByVal lane As Long) As Double

    Select Case lane

        Case 1
            Topo2_InterpolateLane = axisA + ((axisB - axisA) * 0.25)

        Case 2
            Topo2_InterpolateLane = axisA + ((axisB - axisA) * 0.5)

        Case 3
            Topo2_InterpolateLane = axisA + ((axisB - axisA) * 0.75)

        Case Else
            Err.Raise vbObjectError + 5070, "Topo2_InterpolateLane", _
                      "Lane must be 1, 2, or 3."

    End Select

End Function


Private Function Topo2_ApplyHorizontalGridStep( _
    ByVal gridCol As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String) As Long

    Select Case UCase$(directionCode)

        Case "L"
            Topo2_ApplyHorizontalGridStep = gridCol - steps

        Case "R"
            Topo2_ApplyHorizontalGridStep = gridCol + steps

        Case Else
            Err.Raise vbObjectError + 5080, "Topo2_ApplyHorizontalGridStep", _
                      "Direction must be L or R."

    End Select

End Function


Private Function Topo2_ApplyVerticalGridStep( _
    ByVal gridRow As Long, _
    ByVal steps As Long, _
    ByVal directionCode As String) As Long

    Select Case UCase$(directionCode)

        Case "U"
            Topo2_ApplyVerticalGridStep = gridRow - steps

        Case "D"
            Topo2_ApplyVerticalGridStep = gridRow + steps

        Case Else
            Err.Raise vbObjectError + 5090, "Topo2_ApplyVerticalGridStep", _
                      "Direction must be U or D."

    End Select

End Function


Private Function Topo2_GetAxisValue( _
    ByVal axisValues As Variant, _
    ByVal gridIndex As Long) As Double

    Topo2_GetAxisValue = CDbl(axisValues(gridIndex - 1))

End Function


' ============================================================
' Output and validation helpers
' ============================================================

Private Function Topo2_CreatePoint( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal vertexType As String, _
    ByVal label As String) As Object

    Dim point As Object

    Set point = CreateObject("Scripting.Dictionary")
    point.CompareMode = vbTextCompare

    point("x") = x
    point("y") = y
    point("vertextype") = vertexType
    point("label") = label

    Set Topo2_CreatePoint = point

End Function


Private Sub Topo2_PrintRelationshipChecks( _
    ByVal routeName As String, _
    ByVal points As Collection)

    Dim i As Long

    Debug.Print "RELATIONSHIP CHECKS FOR " & routeName

    For i = 1 To points.Count - 1

        Debug.Print _
            "P" & CStr(i) & " -> P" & CStr(i + 1) & _
            " share axis: " & _
            CStr(Topo2_ShareAxis(points(i), points(i + 1)))

    Next i

End Sub


Private Function Topo2_ShareAxis( _
    ByVal p1 As Object, _
    ByVal p2 As Object) As Boolean

    Topo2_ShareAxis = _
        (Abs(CDbl(p1("x")) - CDbl(p2("x"))) < 0.001) Or _
        (Abs(CDbl(p1("y")) - CDbl(p2("y"))) < 0.001)

End Function


Private Sub Topo2_ValidateLane( _
    ByVal lane As Long, _
    ByVal laneType As String, _
    ByVal sourceText As String)

    If lane < 1 Or lane > 3 Then
        Err.Raise vbObjectError + 5100, "Topo2_ValidateLane", _
                  laneType & " must be 1, 2, or 3. Source: " & sourceText
    End If

End Sub


Private Function Topo2_MinLong(ByVal a As Long, ByVal b As Long) As Long

    If a < b Then
        Topo2_MinLong = a
    Else
        Topo2_MinLong = b
    End If

End Function


Private Function Topo2_MaxLong(ByVal a As Long, ByVal b As Long) As Long

    If a > b Then
        Topo2_MaxLong = a
    Else
        Topo2_MaxLong = b
    End If

End Function


Private Function Topo2_Num(ByVal value As Double) As String

    Topo2_Num = Replace(CStr(value), Application.International(xlDecimalSeparator), ".")

End Function