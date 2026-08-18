Public Function ParseShapeRow(ByVal csvLine As String, ByVal keyToIndex As Object) As Object

    Dim fields As Variant
    Dim shape As Object
    Dim rowType As String

    fields = ParseCsvLine(csvLine)

    Set shape = CreateObject("Scripting.Dictionary")
    shape.CompareMode = vbTextCompare

    shape(KEY_ID) = GetFieldValue(fields, keyToIndex, KEY_ID)
    shape(KEY_TYPE) = GetFieldValue(fields, keyToIndex, KEY_TYPE)

    rowType = CStr(shape(KEY_TYPE))

    shape(KEY_ROUTEDEF) = GetFieldValue( _
        fields, _
        keyToIndex, _
        KEY_ROUTEDEF _
    )

    If UCase$(rowType) = "_META" Then

        shape(KEY_PAGEWIDTH) = GetFieldValue(fields, keyToIndex, KEY_PAGEWIDTH)
        shape(KEY_PAGEHEIGHT) = GetFieldValue(fields, keyToIndex, KEY_PAGEHEIGHT)
        shape(KEY_PAGETITLE) = GetFieldValue(fields, keyToIndex, KEY_PAGETITLE)
        shape(KEY_BACKGROUND) = GetFieldValue(fields, keyToIndex, KEY_BACKGROUND)
        shape(KEY_TITLECOLOUR) = GetFieldValue(fields, keyToIndex, KEY_TITLECOLOUR)

        shape(KEY_HGRID) = GetFieldValue(fields, keyToIndex, KEY_HGRID)
        shape(KEY_VGRID) = GetFieldValue(fields, keyToIndex, KEY_VGRID)

        shape(KEY_ROUTECOLOUR) = GetFieldValue(fields, keyToIndex, KEY_ROUTECOLOUR)
        shape(KEY_ROUTEWIDTH) = GetFieldValue(fields, keyToIndex, KEY_ROUTEWIDTH)
        shape(KEY_ROUTEARROWSIZE) = GetFieldValue(fields, keyToIndex, KEY_ROUTEARROWSIZE)
        shape(KEY_ROUTEARROWTYPE) = GetFieldValue(fields, keyToIndex, KEY_ROUTEARROWTYPE)

        shape(KEY_MASKOFFSET) = GetFieldValue(fields, keyToIndex, KEY_MASKOFFSET)
        
    ElseIf LCase$(rowType) = "route" Then

        shape(KEY_ROUTEDEF) = GetFieldValue( _
            fields, _
            keyToIndex, _
            KEY_ROUTEDEF _
        )

        shape(KEY_ROUTECOLOUR) = GetFieldValue( _
            fields, _
            keyToIndex, _
            KEY_ROUTECOLOUR _
        )

        shape(KEY_ROUTEWIDTH) = GetFieldValue( _
            fields, _
            keyToIndex, _
            KEY_ROUTEWIDTH _
        )

        shape(KEY_ROUTEARROWSIZE) = GetFieldValue( _
            fields, _
            keyToIndex, _
            KEY_ROUTEARROWSIZE _
        )

        shape(KEY_ROUTEARROWTYPE) = GetFieldValue( _
            fields, _
            keyToIndex, _
            KEY_ROUTEARROWTYPE _
        )
        
    Else
        shape(KEY_GRIDCOL) = CLng(GetFieldValue(fields, keyToIndex, KEY_GRIDCOL))
        shape(KEY_GRIDROW) = CLng(GetFieldValue(fields, keyToIndex, KEY_GRIDROW))
        shape(KEY_SIZEX) = ToDbl(GetFieldValue(fields, keyToIndex, KEY_SIZEX))
        shape(KEY_SIZEY) = ToDbl(GetFieldValue(fields, keyToIndex, KEY_SIZEY))
        shape(KEY_RAD) = ToDbl(GetFieldValue(fields, keyToIndex, KEY_RAD))

        shape(KEY_FILLCOL) = GetFieldValue(fields, keyToIndex, KEY_FILLCOL)
        shape(KEY_STROKECOL) = GetFieldValue(fields, keyToIndex, KEY_STROKECOL)
        shape(KEY_TEXTCOL) = GetFieldValue(fields, keyToIndex, KEY_TEXTCOL)

        shape(KEY_CAPTION) = GetFieldValue(fields, keyToIndex, KEY_CAPTION)
        shape(KEY_TOOLTIP) = GetFieldValue(fields, keyToIndex, KEY_TOOLTIP)
        shape(KEY_LINKURL) = GetFieldValue(fields, keyToIndex, KEY_LINKURL)
        shape(KEY_NEWPAGE) = GetFieldValue(fields, keyToIndex, KEY_NEWPAGE)

    End If

    Set ParseShapeRow = shape

End Function


