Option Explicit

' ============================================================
' Simplified routing engine
'
' Purpose:
'   - Build SHAPEGRID from CSV _META hgrid/vgrid
'   - Derive CORRIDORGRID between SHAPEGRID axes
'   - Merge both into FULLGRID
'   - Parse user-authored route strings
'   - Output route coordinates to the Immediate Window
'
' Current stage:
'   - Debug output only
'   - No SVG line creation yet
'   - No topology intelligence
'   - No route validation beyond basic parsing/failure reporting
'
' Route format:
'
'   start | optional movements | end
'
' Start token:
'   shapeID-exitSocket-toCorridorLane
'
'   Example:
'       S22-E3-2
'
' Movement token:
'   moveDirectionAndDistance-toLane
'
'   Examples:
'       D3-1
'       L1-3
'
' End token:
'   shapeID-entrySocket
'
'   Example:
'       S41-N1
'
' Lane convention:
'   Lane 1 = axis - ROUTE_LANE_SEP
'   Lane 2 = axis
'   Lane 3 = axis + ROUTE_LANE_SEP
'
' Socket convention:
'   N1/N2/N3 are placed across the top edge at 1/4, 1/2, 3/4 width
'   S1/S2/S3 are placed across the bottom edge at 1/4, 1/2, 3/4 width
'   W1/W2/W3 are placed down the left edge at 1/4, 1/2, 3/4 height
'   E1/E2/E3 are placed down the right edge at 1/4, 1/2, 3/4 height
' ============================================================

Private Const ROUTE_LANE_SEP As Double = 15
Private Const EPS As Double = 0.000001

Private Const AXIS_VALUE As String = "value"
Private Const AXIS_KIND As String = "kind"

Private Const GRID_SHAPE As String = "SHAPE"
Private Const GRID_CORRIDOR As String = "CORRIDOR"

Private Const POINT_TYPE As String = "type"
Private Const POINT_LABEL As String = "label"
Private Const POINT_X As String = "x"
Private Const POINT_Y As String = "y"


' ============================================================
' Public debug entry point
' ============================================================

Public Sub DebugPrintRoutes( _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object, _
    ByVal routeStrings As Collection)

    Dim fullGrid As Object
    Dim shapeIndex As Object
    Dim routeText As Variant
    Dim routePoints As Collection

    Set fullGrid = BuildFullGrid(pageSettings)
    Set shapeIndex = BuildShapeIndex(shapes)

    Debug.Print String(80, "=")
    Debug.Print "ROUTING DEBUG"
    Debug.Print String(80, "=")

    DebugPrintFullGrid fullGrid

    For Each routeText In routeStrings

        Debug.Print vbCrLf & String(80, "-")
        Debug.Print "ROUTE: " & CStr(routeText)
        Debug.Print String(80, "-")

        On Error GoTo RouteFailed

        Set routePoints = ParseRouteToPoints( _
            CStr(routeText), _
            shapeIndex, _
            fullGrid _
        )

        DebugPrintRoutePoints routePoints

        On Error GoTo 0

ContinueRouteLoop:
    Next routeText

    Debug.Print String(80, "=")
    Debug.Print "ROUTING DEBUG COMPLETE"
    Debug.Print String(80, "=")

    Exit Sub

RouteFailed:
    Debug.Print "ROUTE FAILED: " & CStr(routeText)
    Debug.Print "ERROR: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Resume ContinueRouteLoop

End Sub

Public Sub DebugPrintRoutesFromCsv( _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object, _
    ByVal routes As Collection)

    Dim routeStrings As Collection
    Dim routeRow As Object

    Set routeStrings = New Collection

    For Each routeRow In routes

        If Len(Trim$(CStr(routeRow(KEY_ROUTEDEF)))) > 0 Then

            routeStrings.Add _
                CStr(routeRow(KEY_ROUTEDEF))

        End If

    Next routeRow

    DebugPrintRoutes _
        shapes, _
        pageSettings, _
        routeStrings

End Sub

Public Function BuildRoutesSvg( _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object, _
    ByVal routes As Collection) As String

    Dim sb As String
    Dim fullGrid As Object
    Dim shapeIndex As Object

    Dim routeRow As Object
    Dim routeDef As String
    Dim routeId As String
    Dim routePoints As Collection

    If routes Is Nothing Then
        BuildRoutesSvg = vbNullString
        Exit Function
    End If

    If routes.Count = 0 Then
        BuildRoutesSvg = vbNullString
        Exit Function
    End If

    Set fullGrid = BuildFullGrid(pageSettings)
    Set shapeIndex = BuildShapeIndex(shapes)

    sb = sb & "  <!-- ROUTES -->" & vbCrLf

    sb = sb & " <defs>" & vbCrLf

    For Each routeRow In routes

        routeId = CStr(routeRow(KEY_ID))
        routeDef = Trim$(CStr(routeRow(KEY_ROUTEDEF)))

        If Len(routeDef) > 0 Then

            sb = sb & BuildRouteMarkerSvg( _
                routeRow, _
                pageSettings)

        End If

    Next routeRow

    sb = sb & " </defs>" & vbCrLf

    For Each routeRow In routes

        routeId = CStr(routeRow(KEY_ID))
        routeDef = Trim$(CStr(routeRow(KEY_ROUTEDEF)))

        If Len(routeDef) > 0 Then

            On Error GoTo RouteSvgFailed

            Set routePoints = ParseRouteToPoints( _
                routeDef, _
                shapeIndex, _
                fullGrid _
            )

            sb = sb & BuildRoutePolylineSvg( _
                routeRow, _
                pageSettings, _
                routePoints _
            )

            On Error GoTo 0

        End If

    ContinueRouteLoop:
    Next routeRow

    sb = sb & "  <!-- END ROUTES -->" & vbCrLf

    BuildRoutesSvg = sb

    Exit Function

RouteSvgFailed:

    Err.Raise vbObjectError + 3600, _
              "BuildRoutesSvg", _
              "Failed building SVG route '" & routeId & _
              "' with definition '" & routeDef & "'. " & _
              Err.Description

End Function

Private Function BuildRoutePolylineSvg( _
    ByVal routeRow As Object, _
    ByVal pageSettings As Object, _
    ByVal routePoints As Collection) As String

    Dim sb As String

    Dim routeId As String
    Dim routeDef As String

    Dim routeColour As String
    Dim routeWidth As Double
    Dim routeArrowType As String

    Dim markerStart As String
    Dim markerEnd As String

    Dim pointsText As String

    Dim i As Long
    Dim point As Object

    routeId = CStr(routeRow(KEY_ID))
    routeDef = CStr(routeRow(KEY_ROUTEDEF))

    routeColour = ResolveStringValue( _
        routeRow, _
        pageSettings, _
        KEY_ROUTECOLOUR, _
        "#000000")

    routeWidth = ResolveDoubleValue( _
        routeRow, _
        pageSettings, _
        KEY_ROUTEWIDTH, _
        4)

    routeArrowType = LCase$(ResolveStringValue( _
        routeRow, _
        pageSettings, _
        KEY_ROUTEARROWTYPE, _
        "e"))

    Select Case routeArrowType

        Case "s"

            markerStart = _
                " marker-start=""url(#routeArrowReverse_" & routeId & ")"" "

            markerEnd = vbNullString

        Case "e"

            markerStart = vbNullString

            markerEnd = _
                " marker-end=""url(#routeArrow_" & routeId & ")"" "

        Case "b"

            markerStart = _
                " marker-start=""url(#routeArrowReverse_" & routeId & ")"" "

            markerEnd = _
                " marker-end=""url(#routeArrow_" & routeId & ")"" "


        Case Else

            markerStart = vbNullString
            markerEnd = vbNullString

    End Select

    For i = 1 To routePoints.Count

        Set point = routePoints(i)

        If Len(pointsText) > 0 Then
            pointsText = pointsText & " "
        End If

        pointsText = pointsText & _
            SvgNum(CDbl(point(POINT_X))) & "," & _
            SvgNum(CDbl(point(POINT_Y)))

    Next i

    sb = sb & "  <polyline" & _
              " id=""route-" & XmlSafeAttribute(routeId) & """" & _
              " points=""" & pointsText & """" & _
              " fill=""none""" & _
              " stroke=""" & routeColour & """" & _
              " stroke-width=""" & SvgNum(routeWidth) & """" & _
              markerStart & _
              markerEnd & _
              " stroke-linecap=""round""" & _
              " stroke-linejoin=""round""" & _
              ">" & vbCrLf

    sb = sb & "    <title>" & _
              XmlSafeText(routeId & ": " & routeDef) & _
              "</title>" & vbCrLf

    sb = sb & "  </polyline>" & vbCrLf

    BuildRoutePolylineSvg = sb

End Function

Private Function BuildRouteMarkerSvg( _
    ByVal routeRow As Object, _
    ByVal pageSettings As Object) As String

    Dim sb As String

    Dim routeId As String
    Dim routeColour As String

    Dim arrowSize As Double
    Dim halfArrow As Double
    Dim arrowLength As Double

    routeId = CStr(routeRow(KEY_ID))

    routeColour = ResolveStringValue( _
        routeRow, _
        pageSettings, _
        KEY_ROUTECOLOUR, _
        "#000000")

    arrowSize = ResolveDoubleValue( _
        routeRow, _
        pageSettings, _
        KEY_ROUTEARROWSIZE, _
        6)

    halfArrow = arrowSize / 2

    ' Arrow size represents the pixel width of the arrow base.
    ' Arrow length is 1.5 × the base width.

    arrowLength = arrowSize * 1.5

    ' --------------------------------------------------
    ' Forward arrow
    ' --------------------------------------------------

    sb = sb & "    <marker" & _
            " id=""routeArrow_" & routeId & """" & _
            " markerWidth=""" & SvgNum(arrowLength) & """" & _
            " markerHeight=""" & SvgNum(arrowSize) & """" & _
            " refX=""" & SvgNum(arrowLength) & """" & _
            " refY=""" & SvgNum(halfArrow) & """" & _
            " orient=""auto"">" & vbCrLf

    sb = sb & "      <path" & _
            " d=""M0,0 L" & _
            SvgNum(arrowLength) & "," & _
            SvgNum(halfArrow) & _
            " L0," & _
            SvgNum(arrowSize) & _
            " z""" & _
            " fill=""" & XmlSafeAttribute(routeColour) & """ />" & vbCrLf

    sb = sb & "    </marker>" & vbCrLf

    ' --------------------------------------------------
    ' Reverse arrow
    ' --------------------------------------------------

    sb = sb & "    <marker" & _
            " id=""routeArrowReverse_" & routeId & """" & _
            " markerWidth=""" & SvgNum(arrowLength) & """" & _
            " markerHeight=""" & SvgNum(arrowSize) & """" & _
            " refX=""0""" & _
            " refY=""" & SvgNum(halfArrow) & """" & _
            " orient=""auto"">" & vbCrLf

    sb = sb & "      <path" & _
            " d=""M" & _
            SvgNum(arrowLength) & ",0 L0," & _
            SvgNum(halfArrow) & _
            " L" & _
            SvgNum(arrowLength) & "," & _
            SvgNum(arrowSize) & _
            " z""" & _
            " fill=""" & XmlSafeAttribute(routeColour) & """ />" & vbCrLf

    sb = sb & "    </marker>" & vbCrLf

    BuildRouteMarkerSvg = sb

End Function

Private Function ResolveStringValue( _
    ByVal routeRow As Object, _
    ByVal pageSettings As Object, _
    ByVal key As String, _
    ByVal defaultValue As String) As String

    Dim valueText As String

    valueText = Trim$(CStr(routeRow(key)))

    If Len(valueText) > 0 Then

        ResolveStringValue = valueText
        Exit Function

    End If

    valueText = Trim$(CStr(pageSettings(key)))

    If Len(valueText) > 0 Then

        ResolveStringValue = valueText
        Exit Function

    End If

    ResolveStringValue = defaultValue

End Function


Private Function ResolveDoubleValue( _
    ByVal routeRow As Object, _
    ByVal pageSettings As Object, _
    ByVal key As String, _
    ByVal defaultValue As Double) As Double

    Dim valueText As String

    valueText = Trim$(CStr(routeRow(key)))

    If Len(valueText) > 0 Then

        ResolveDoubleValue = CDbl(valueText)
        Exit Function

    End If

    valueText = Trim$(CStr(pageSettings(key)))

    If Len(valueText) > 0 Then

        ResolveDoubleValue = CDbl(valueText)
        Exit Function

    End If

    ResolveDoubleValue = defaultValue

End Function


' ============================================================
' Phase 1: Build SHAPEGRID, CORRIDORGRID and FULLGRID
' ============================================================

Private Function BuildFullGrid(ByVal pageSettings As Object) As Object

    Dim result As Object

    Dim shapeX As Collection
    Dim shapeY As Collection

    Dim corridorX As Collection
    Dim corridorY As Collection

    Dim fullX As Collection
    Dim fullY As Collection

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Set shapeX = ParseGridText(CStr(pageSettings(KEY_HGRID)))
    Set shapeY = ParseGridText(CStr(pageSettings(KEY_VGRID)))

    Set corridorX = BuildCorridorAxes(shapeX)
    Set corridorY = BuildCorridorAxes(shapeY)

    Set fullX = MergeAxes(shapeX, corridorX)
    Set fullY = MergeAxes(shapeY, corridorY)

    result.Add "ShapeX", shapeX
    result.Add "ShapeY", shapeY

    result.Add "CorridorX", corridorX
    result.Add "CorridorY", corridorY

    result.Add "FullX", fullX
    result.Add "FullY", fullY

    Set BuildFullGrid = result

End Function


Private Function ParseGridText(ByVal gridText As String) As Collection

    Dim result As New Collection
    Dim parts As Variant
    Dim i As Long
    Dim valueText As String

    gridText = Replace(gridText, """", "")
    gridText = Trim$(gridText)

    If Len(gridText) = 0 Then
        Err.Raise vbObjectError + 3000, _
                  "ParseGridText", _
                  "Grid text is empty."
    End If

    parts = Split(gridText, "|")

    For i = LBound(parts) To UBound(parts)
        valueText = Trim$(CStr(parts(i)))

        If Len(valueText) > 0 Then
            result.Add CDbl(valueText)
        End If
    Next i

    If result.Count = 0 Then
        Err.Raise vbObjectError + 3001, _
                  "ParseGridText", _
                  "Grid text did not contain any numeric values."
    End If

    Set ParseGridText = SortAxisValues(result)

End Function


Private Function BuildCorridorAxes(ByVal shapeAxes As Collection) As Collection

    Dim result As New Collection
    Dim i As Long

    If shapeAxes.Count < 2 Then
        Set BuildCorridorAxes = result
        Exit Function
    End If

    For i = 1 To shapeAxes.Count - 1
        result.Add (CDbl(shapeAxes(i)) + CDbl(shapeAxes(i + 1))) / 2
    Next i

    Set BuildCorridorAxes = result

End Function


Private Function MergeAxes( _
    ByVal shapeAxes As Collection, _
    ByVal corridorAxes As Collection) As Collection

    Dim mergedValues As New Collection
    Dim result As Collection
    Dim v As Variant

    For Each v In shapeAxes
        mergedValues.Add CDbl(v)
    Next v

    For Each v In corridorAxes
        mergedValues.Add CDbl(v)
    Next v

    Set result = SortAxisValues(mergedValues)

    Set MergeAxes = result

End Function


Private Function SortAxisValues(ByVal values As Collection) As Collection

    Dim arr() As Double
    Dim i As Long
    Dim j As Long
    Dim temp As Double
    Dim result As New Collection

    If values.Count = 0 Then
        Set SortAxisValues = result
        Exit Function
    End If

    ReDim arr(1 To values.Count)

    For i = 1 To values.Count
        arr(i) = CDbl(values(i))
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    For i = LBound(arr) To UBound(arr)
        result.Add arr(i)
    Next i

    Set SortAxisValues = result

End Function


Private Sub DebugPrintFullGrid(ByVal fullGrid As Object)

    Debug.Print vbCrLf & "SHAPEGRID X:"
    DebugPrintAxisCollection fullGrid("ShapeX")

    Debug.Print vbCrLf & "SHAPEGRID Y:"
    DebugPrintAxisCollection fullGrid("ShapeY")

    Debug.Print vbCrLf & "CORRIDORGRID X:"
    DebugPrintAxisCollection fullGrid("CorridorX")

    Debug.Print vbCrLf & "CORRIDORGRID Y:"
    DebugPrintAxisCollection fullGrid("CorridorY")

    Debug.Print vbCrLf & "FULLGRID X:"
    DebugPrintAxisCollection fullGrid("FullX")

    Debug.Print vbCrLf & "FULLGRID Y:"
    DebugPrintAxisCollection fullGrid("FullY")

End Sub


Private Sub DebugPrintAxisCollection(ByVal axes As Collection)

    Dim i As Long

    If axes.Count = 0 Then
        Debug.Print "  <none>"
        Exit Sub
    End If

    For i = 1 To axes.Count
        Debug.Print "  " & CStr(i) & ": " & SvgNum(CDbl(axes(i)))
    Next i

End Sub


' ============================================================
' Phase 2: Build shape lookup
' ============================================================

Private Function BuildShapeIndex(ByVal shapes As Collection) As Object

    Dim result As Object
    Dim shape As Object
    Dim shapeId As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For Each shape In shapes

        shapeId = CStr(shape(KEY_ID))

        If Len(Trim$(shapeId)) > 0 Then
            If Not result.Exists(shapeId) Then
                result.Add shapeId, shape
            End If
        End If

    Next shape

    Set BuildShapeIndex = result

End Function


Private Function GetShapeById( _
    ByVal shapeIndex As Object, _
    ByVal shapeId As String) As Object

    If Not shapeIndex.Exists(shapeId) Then
        Err.Raise vbObjectError + 3100, _
                  "GetShapeById", _
                  "Shape ID not found: " & shapeId
    End If

    Set GetShapeById = shapeIndex(shapeId)

End Function


' ============================================================
' Phase 3: Socket coordinate resolver
' ============================================================

Private Function GetSocketPoint( _
    ByVal shape As Object, _
    ByVal socketText As String) As Object

    Dim sideCode As String
    Dim socketIndex As Long

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    Dim px As Double
    Dim py As Double

    socketText = UCase$(Trim$(socketText))

    If Len(socketText) < 2 Then
        Err.Raise vbObjectError + 3200, _
                  "GetSocketPoint", _
                  "Invalid socket text: " & socketText
    End If

    sideCode = Left$(socketText, 1)
    socketIndex = CLng(Mid$(socketText, 2))

    If socketIndex < 1 Or socketIndex > 3 Then
        Err.Raise vbObjectError + 3201, _
                  "GetSocketPoint", _
                  "Socket index must be 1, 2 or 3: " & socketText
    End If

    x = CDbl(shape(KEY_POSX))
    y = CDbl(shape(KEY_POSY))
    w = CDbl(shape(KEY_SIZEX))
    h = CDbl(shape(KEY_SIZEY))

    Dim centreX as Double
    Dim centreY as Double

    centreX = x + (w/2)
    centreY = y + (h/2)


    Select Case sideCode

        Case "N"

            px = centreX + SocketOffset(socketIndex)
            py = y

        Case "S"

            px = centreX + SocketOffset(socketIndex)
            py = y + h

        Case "W"

            px = x
            py = centreY + SocketOffset(socketIndex)

        Case "E"

            px = x + w
            py = centreY + SocketOffset(socketIndex)

    End Select

    Set GetSocketPoint = CreatePoint("SOCKET", socketText, px, py)

End Function

Private Function FindNearestSocket( _
    ByVal currentPoint As Object, _
    ByVal shape As Object) As Object

    Dim socketList As Variant
    Dim socketName As Variant

    Dim candidate As Object

    Dim bestPoint As Object
    Dim bestDistance As Double
    Dim thisDistance As Double

    socketList = Array( _
        "N1", "N2", "N3", _
        "E1", "E2", "E3", _
        "S1", "S2", "S3", _
        "W1", "W2", "W3")

    bestDistance = 1E+30

    For Each socketName In socketList

        Set candidate = GetSocketPoint( _
            shape, _
            CStr(socketName))

        thisDistance = DistanceBetweenPoints( _
            CDbl(currentPoint(POINT_X)), _
            CDbl(currentPoint(POINT_Y)), _
            CDbl(candidate(POINT_X)), _
            CDbl(candidate(POINT_Y)) _
        )

        If thisDistance < bestDistance Then

            bestDistance = thisDistance

            Set bestPoint = candidate

        End If

    Next socketName

    Set FindNearestSocket = bestPoint

End Function

Private Function DistanceBetweenPoints( _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double) As Double

    DistanceBetweenPoints = _
        Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

End Function



Private Function SocketOffset(ByVal socketIndex As Long) As Double

    Select Case socketIndex

        Case 1
            SocketOffset = -15

        Case 2
            SocketOffset = 0

        Case 3
            SocketOffset = 15

        Case Else

            Err.Raise vbObjectError + 3210, _
                      "SocketOffset", _
                      "Socket index must be 1, 2 or 3."

    End Select

End Function


' ============================================================
' Phase 4: Lane resolver
' ============================================================

Private Function LaneCoordinate( _
    ByVal axisValue As Double, _
    ByVal laneNo As Long) As Double

    Select Case laneNo

        Case 1
            LaneCoordinate = axisValue - ROUTE_LANE_SEP

        Case 2
            LaneCoordinate = axisValue

        Case 3
            LaneCoordinate = axisValue + ROUTE_LANE_SEP

        Case Else
            Err.Raise vbObjectError + 3300, _
                      "LaneCoordinate", _
                      "Lane number must be 1, 2 or 3."

    End Select

End Function


Private Function GetNextCorridorAxis( _
    ByVal corridorAxes As Collection, _
    ByVal currentCoord As Double, _
    ByVal directionCode As String) As Double

    Dim i As Long

    directionCode = UCase$(directionCode)

    Select Case directionCode

        Case "E", "S"

            For i = 1 To corridorAxes.Count
                If CDbl(corridorAxes(i)) > currentCoord + EPS Then
                    GetNextCorridorAxis = CDbl(corridorAxes(i))
                    Exit Function
                End If
            Next i

        Case "W", "N"

            For i = corridorAxes.Count To 1 Step -1
                If CDbl(corridorAxes(i)) < currentCoord - EPS Then
                    GetNextCorridorAxis = CDbl(corridorAxes(i))
                    Exit Function
                End If
            Next i

        Case Else
            Err.Raise vbObjectError + 3310, _
                      "GetNextCorridorAxis", _
                      "Direction must be N, S, E or W."

    End Select

    Err.Raise vbObjectError + 3311, _
              "GetNextCorridorAxis", _
              "No corridor axis found in direction " & directionCode & _
              " from coordinate " & SvgNum(currentCoord)

End Function


Private Function GetNthFullGridAxis( _
    ByVal fullAxes As Collection, _
    ByVal currentCoord As Double, _
    ByVal directionCode As String, _
    ByVal distanceCount As Long) As Double

    Dim i As Long
    Dim hitCount As Long

    If distanceCount < 1 Then
        Err.Raise vbObjectError + 3320, _
                  "GetNthFullGridAxis", _
                  "Movement distance must be 1 or greater."
    End If

    directionCode = UCase$(directionCode)

    Select Case directionCode

        Case "R", "D"

            For i = 1 To fullAxes.Count

                If CDbl(fullAxes(i)) > currentCoord + EPS Then
                    hitCount = hitCount + 1

                    If hitCount = distanceCount Then
                        GetNthFullGridAxis = CDbl(fullAxes(i))
                        Exit Function
                    End If
                End If

            Next i

        Case "L", "U"

            For i = fullAxes.Count To 1 Step -1

                If CDbl(fullAxes(i)) < currentCoord - EPS Then
                    hitCount = hitCount + 1

                    If hitCount = distanceCount Then
                        GetNthFullGridAxis = CDbl(fullAxes(i))
                        Exit Function
                    End If
                End If

            Next i

        Case Else
            Err.Raise vbObjectError + 3321, _
                      "GetNthFullGridAxis", _
                      "Movement direction must be U, D, L or R."

    End Select

    Err.Raise vbObjectError + 3322, _
              "GetNthFullGridAxis", _
              "Unable to find full-grid axis " & CStr(distanceCount) & _
              " in direction " & directionCode & _
              " from coordinate " & SvgNum(currentCoord)

End Function

Private Function NearestFullGridAxis( _
    ByVal coordinate As Double, _
    ByVal fullAxes As Collection) As Double

    Dim i As Long
    Dim bestAxis As Double
    Dim bestDistance As Double
    Dim thisDistance As Double

    bestDistance = 1E+30

    For i = 1 To fullAxes.Count

        thisDistance = Abs( _
            CDbl(fullAxes(i)) - coordinate)

        If thisDistance < bestDistance Then
            bestDistance = thisDistance
            bestAxis = CDbl(fullAxes(i))
        End If

    Next i

    NearestFullGridAxis = bestAxis

End Function

' ============================================================
' Phase 5: Route parser
' ============================================================

Private Function ParseRouteToPoints( _
    ByVal routeText As String, _
    ByVal shapeIndex As Object, _
    ByVal fullGrid As Object) As Collection

    Dim result As New Collection
    Dim tokens As Variant
    Dim i As Long

    Dim startToken As String
    Dim endToken As String
    Dim movementToken As String

    Dim currentPoint As Object

    routeText = Trim$(routeText)

    If Len(routeText) = 0 Then
        Err.Raise vbObjectError + 3400, _
                  "ParseRouteToPoints", _
                  "Route text is empty."
    End If

    tokens = Split(routeText, "|")

    If UBound(tokens) - LBound(tokens) + 1 < 2 Then
        Err.Raise vbObjectError + 3401, _
                  "ParseRouteToPoints", _
                  "Route must contain at least start and end tokens."
    End If

    startToken = Trim$(CStr(tokens(LBound(tokens))))
    endToken = Trim$(CStr(tokens(UBound(tokens))))

    Set currentPoint = ResolveStartToken(startToken, shapeIndex, fullGrid, result)

    If UBound(tokens) > LBound(tokens) + 1 Then

        For i = LBound(tokens) + 1 To UBound(tokens) - 1

            movementToken = Trim$(CStr(tokens(i)))

            Set currentPoint = ResolveMovementToken( _
                movementToken, _
                currentPoint, _
                fullGrid, _
                result _
            )

        Next i

    End If

    ResolveEndToken endToken, currentPoint, shapeIndex, result

    Set ParseRouteToPoints = result

End Function


Private Function ResolveStartToken( _
    ByVal startToken As String, _
    ByVal shapeIndex As Object, _
    ByVal fullGrid As Object, _
    ByVal routePoints As Collection) As Object

    Dim parts As Variant

    Dim shapeId As String
    Dim socketText As String
    Dim sideCode As String
    Dim laneNo As Long

    Dim shape As Object
    Dim socketPoint As Object
    Dim corridorPoint As Object

    Dim targetAxis As Double
    Dim targetX As Double
    Dim targetY As Double

    parts = Split(startToken, "-")

    If UBound(parts) - LBound(parts) + 1 <> 3 Then
        Err.Raise vbObjectError + 3410, _
                  "ResolveStartToken", _
                  "Start token must be shapeID-socket-lane: " & startToken
    End If

    shapeId = Trim$(CStr(parts(0)))
    socketText = UCase$(Trim$(CStr(parts(1))))
    laneNo = CLng(Trim$(CStr(parts(2))))

    sideCode = Left$(socketText, 1)

    Set shape = GetShapeById(shapeIndex, shapeId)
    Set socketPoint = GetSocketPoint(shape, socketText)

    socketPoint(POINT_TYPE) = "START_SOCKET"
    socketPoint(POINT_LABEL) = startToken

    routePoints.Add socketPoint

    targetX = CDbl(socketPoint(POINT_X))
    targetY = CDbl(socketPoint(POINT_Y))

    Select Case sideCode

        Case "E", "W"

            targetAxis = GetNextCorridorAxis( _
                fullGrid("CorridorX"), _
                CDbl(socketPoint(POINT_X)), _
                sideCode _
            )

            targetX = LaneCoordinate(targetAxis, laneNo)

        Case "N", "S"

            targetAxis = GetNextCorridorAxis( _
                fullGrid("CorridorY"), _
                CDbl(socketPoint(POINT_Y)), _
                sideCode _
            )

            targetY = LaneCoordinate(targetAxis, laneNo)

        Case Else
            Err.Raise vbObjectError + 3411, _
                      "ResolveStartToken", _
                      "Start socket side must be N, S, E or W: " & socketText

    End Select

    Set corridorPoint = CreatePoint( _
        "START_CORRIDOR", _
        startToken, _
        targetX, _
        targetY _
    )

    routePoints.Add corridorPoint

    Set ResolveStartToken = corridorPoint

End Function


Private Function ResolveMovementToken( _
    ByVal movementToken As String, _
    ByVal currentPoint As Object, _
    ByVal fullGrid As Object, _
    ByVal routePoints As Collection) As Object

    Dim parts As Variant
    Dim moveText As String
    Dim directionCode As String
    Dim distanceCount As Long
    Dim laneNo As Long

    Dim targetAxis As Double
    Dim targetX As Double
    Dim targetY As Double

    parts = Split(movementToken, "-")

    If UBound(parts) - LBound(parts) + 1 <> 2 Then
        Err.Raise vbObjectError + 3420, _
                  "ResolveMovementToken", _
                  "Movement token must be directionDistance-lane: " & movementToken
    End If

    moveText = UCase$(Trim$(CStr(parts(0))))
    laneNo = CLng(Trim$(CStr(parts(1))))

    If Len(moveText) < 2 Then
        Err.Raise vbObjectError + 3421, _
                  "ResolveMovementToken", _
                  "Movement must contain direction and distance: " & movementToken
    End If

    directionCode = Left$(moveText, 1)
    distanceCount = CLng(Mid$(moveText, 2))

    targetX = CDbl(currentPoint(POINT_X))
    targetY = CDbl(currentPoint(POINT_Y))

    Dim currentAxisX As Double
    Dim currentAxisY As Double

    currentAxisX = NearestFullGridAxis( _
        targetX, _
        fullGrid("FullX"))

    currentAxisY = NearestFullGridAxis( _
        targetY, _
        fullGrid("FullY"))
   

    Select Case directionCode

        Case "R", "L"

            targetAxis = GetNthFullGridAxis( _
                fullGrid("FullX"), _
                currentAxisX, _
                directionCode, _
                distanceCount _
            )

            targetX = LaneCoordinate(targetAxis, laneNo)

        Case "D", "U"

            targetAxis = GetNthFullGridAxis( _
                fullGrid("FullY"), _
                currentAxisY, _
                directionCode, _
                distanceCount _
            )

            targetY = LaneCoordinate(targetAxis, laneNo)

        Case Else
            Err.Raise vbObjectError + 3422, _
                      "ResolveMovementToken", _
                      "Movement direction must be U, D, L or R: " & movementToken

    End Select

    Set ResolveMovementToken = CreatePoint( _
        "MOVE_" & directionCode, _
        movementToken, _
        targetX, _
        targetY _
    )

    routePoints.Add ResolveMovementToken

End Function


Private Sub ResolveEndToken( _
    ByVal endToken As String, _
    ByVal currentPoint As Object, _
    ByVal shapeIndex As Object, _
    ByVal routePoints As Collection)

    Dim shape As Object
    Dim endPoint As Object

    Set shape = GetShapeById(shapeIndex, endToken)

    Set endPoint = FindNearestSocket( _
        currentPoint, _
        shape _
    )

    endPoint(POINT_TYPE) = "END_SOCKET"
    endPoint(POINT_LABEL) = endToken

    routePoints.Add endPoint

End Sub



' ============================================================
' Phase 6: Immediate Window output helpers
' ============================================================

Private Function CreatePoint( _
    ByVal pointType As String, _
    ByVal label As String, _
    ByVal x As Double, _
    ByVal y As Double) As Object

    Dim point As Object

    Set point = CreateObject("Scripting.Dictionary")
    point.CompareMode = vbTextCompare

    point.Add POINT_TYPE, pointType
    point.Add POINT_LABEL, label
    point.Add POINT_X, x
    point.Add POINT_Y, y

    Set CreatePoint = point

End Function


Private Sub DebugPrintRoutePoints(ByVal routePoints As Collection)

    Dim i As Long
    Dim point As Object

    If routePoints.Count = 0 Then
        Debug.Print "  <no points>"
        Exit Sub
    End If

    For i = 1 To routePoints.Count

        Set point = routePoints(i)

        Debug.Print CStr(i) & ": " & _
                    CStr(point(POINT_TYPE)) & _
                    "  " & CStr(point(POINT_LABEL)) & _
                    "  x=" & SvgNum(CDbl(point(POINT_X))) & _
                    "  y=" & SvgNum(CDbl(point(POINT_Y)))

    Next i

End Sub