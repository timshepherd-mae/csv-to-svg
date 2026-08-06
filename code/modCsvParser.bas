Public Function ParseCsvLine(ByVal s As String) As Variant

    Dim values As Collection
    Dim currentValue As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String
    Dim inQuotes As Boolean
    Dim arr() As String

    Set values = New Collection
    currentValue = vbNullString
    inQuotes = False

    i = 1

    Do While i <= Len(s)

        ch = Mid$(s, i, 1)

        If ch = """" Then

            If inQuotes And i < Len(s) Then
                nextCh = Mid$(s, i + 1, 1)

                If nextCh = """" Then
                    currentValue = currentValue & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = Not inQuotes
            End If

        ElseIf ch = "," And Not inQuotes Then

            values.Add currentValue
            currentValue = vbNullString

        Else

            currentValue = currentValue & ch

        End If

        i = i + 1

    Loop

    values.Add currentValue

    ReDim arr(0 To values.Count - 1)

    For i = 1 To values.Count
        arr(i - 1) = CStr(values(i))
    Next i

    ParseCsvLine = arr

End Function

Public Function BuildHeaderMap(ByVal headerFields As Variant) As Object

    Dim map As Object
    Dim i As Long
    Dim key As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    For i = LBound(headerFields) To UBound(headerFields)
        key = Trim$(CStr(headerFields(i)))

        If Len(key) > 0 Then
            If Not map.Exists(key) Then
                map.Add key, i
            End If
        End If
    Next i

    Set BuildHeaderMap = map

End Function

Public Sub ValidateRequiredKeys(ByVal keyToIndex As Object)

    Dim requiredKeys As Variant
    Dim i As Long
    Dim missing As String


    requiredKeys = Array( _
        KEY_ID, _
        KEY_TYPE, _
        KEY_POSX, _
        KEY_POSY, _
        KEY_SIZEX, _
        KEY_SIZEY, _
        KEY_RAD, _
        KEY_FILLCOL, _
        KEY_STROKECOL, _
        KEY_TEXTCOL, _
        KEY_CAPTION, _
        KEY_TOOLTIP, _
        KEY_LINKURL, _
        KEY_NEWPAGE, _
        KEY_PAGEWIDTH, _
        KEY_PAGEHEIGHT, _
        KEY_PAGETITLE, _
        KEY_BACKGROUND, _
        KEY_TITLECOLOUR _
    )

    For i = LBound(requiredKeys) To UBound(requiredKeys)
        If Not keyToIndex.Exists(CStr(requiredKeys(i))) Then
            missing = missing & vbCrLf & " - " & CStr(requiredKeys(i))
        End If
    Next i

    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 1100, "ValidateRequiredKeys", _
                  "CSV is missing required key(s):" & missing
    End If

End Sub

Public Function GetFieldValue(ByVal fields As Variant, ByVal keyToIndex As Object, ByVal key As String) As String

    Dim idx As Long

    idx = CLng(keyToIndex(key))

    If idx < LBound(fields) Or idx > UBound(fields) Then
        GetFieldValue = vbNullString
    Else
        GetFieldValue = CStr(fields(idx))
    End If

End Function

