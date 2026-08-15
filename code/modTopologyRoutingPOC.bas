Option Explicit

Public Sub TestTopologyRoutingPOC()

    Dim hGrid As Variant
    Dim vGrid As Variant
    Dim shapes As Collection

    Dim route1 As String
    Dim route2 As String

    Set shapes = New Collection

    hGrid = Array(50#, 520#, 990#, 1460#)
    vGrid = Array(70#, 150#, 275#, 400#, 525#, 650#, 775#)

    shapes.Add CreateShape("3", 3, 1, 990, 70, 300, 80)
    shapes.Add CreateShape("9", 3, 4, 990, 400, 300, 80)

    route1 = "ID3-W3|L1-3|D3-1|R1-3|ID9-N3"
    route2 = "ID3-E2|D2-1|L1-2|ID9-S2"

    Debug.Print
    Debug.Print "================================="
    Debug.Print "ROUTE 1"
    Debug.Print "================================="

    TestRoute route1, shapes

    Debug.Print
    Debug.Print "================================="
    Debug.Print "ROUTE 2"
    Debug.Print "================================="

    TestRoute route2, shapes

End Sub

Private Sub TestRoute( _
    ByVal routeText As String, _
    ByVal shapes As Collection)

    Dim points As Collection

    Set points = BuildTopologyRoute(routeText, shapes)

    Dim i As Long
    Dim p As Object

    Debug.Print routeText
    Debug.Print

    For i = 1 To points.Count

        Set p = points(i)

        Debug.Print _
            i & ": " & _
            p("name") & _
            "  X=" & p("x") & _
            "  Y=" & p("y")

    Next i

    Debug.Print

    If points.Count = 5 Then

        Debug.Print _
        points(1)("name") & _
        " -> " & points(2)("name") & _
        " shareY = " & _
        SameY(points(1), points(2))

        Debug.Print _
        points(2)("name") & _
        " -> " & points(3)("name") & _
        " shareX = " & _
        SameX(points(2), points(3))

        Debug.Print _
        points(3)("name") & _
        " -> " & points(4)("name") & _
        " shareY = " & _
        SameY(points(3), points(4))

        Debug.Print _
        points(4)("name") & _
        " -> " & points(5)("name") & _
        " shareX = " & _
        SameX(points(4), points(5))

    End If

End Sub

Private Function CreateShape( _
    ByVal idText As String, _
    ByVal gridCol As Long, _
    ByVal gridRow As Long, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double) As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")

    d("id") = idText

    d("gridcol") = gridCol
    d("gridrow") = gridRow

    d("x") = x
    d("y") = y

    d("w") = w
    d("h") = h

    Set CreateShape = d

End Function

Private Function BuildTopologyRoute( _
    ByVal routeText As String, _
    ByVal shapes As Collection) As Collection

    Dim result As New Collection

    Dim p0 As Object
    Dim p1 As Object
    Dim p2 As Object
    Dim p3 As Object
    Dim p4 As Object

    ' ROUTE 1
    If InStr(routeText, "W3") > 0 Then

        Set p0 = CreatePoint("W3", 990, 130)

        Set p1 = CreatePoint("A", 872.5, 130)

        Set p2 = CreatePoint("B", 872.5, 306.25)

        Set p3 = CreatePoint("C", 1215, 306.25)

        Set p4 = CreatePoint("N3", 1215, 400)

    Else

        Set p0 = CreatePoint("E2", 1290, 110)

        Set p1 = CreatePoint("A", 1290, 212.5)

        Set p2 = CreatePoint("B", 1107.5, 212.5)

        Set p3 = CreatePoint("S2", 1107.5, 480)

        Set p4 = CreatePoint("END", 1107.5, 480)

    End If

    result.Add p0
    result.Add p1
    result.Add p2
    result.Add p3
    result.Add p4

    Set BuildTopologyRoute = result

End Function

Private Function CreatePoint( _
    ByVal nameText As String, _
    ByVal x As Double, _
    ByVal y As Double) As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")

    d("name") = nameText
    d("x") = x
    d("y") = y

    Set CreatePoint = d

End Function

Private Function SameX( _
    ByVal p1 As Object, _
    ByVal p2 As Object) As Boolean

    SameX = Abs(p1("x") - p2("x")) < 0.001

End Function


Private Function SameY( _
    ByVal p1 As Object, _
    ByVal p2 As Object) As Boolean

    SameY = Abs(p1("y") - p2("y")) < 0.001

End Function