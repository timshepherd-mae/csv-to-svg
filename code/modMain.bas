Option Explicit

' ============================================================
' CSV-to-SVG Conversion Tool
' Stage 01:
'   - Select a CSV file
'   - Parse by hard-coded key names, not by fixed column numbers
'   - Write an SVG file beside the CSV
'   - Use same base file name with .svg extension
'
' Inferred CSV keys from test file:
'   id,posx,posy,sizex,sizey,rad,fillcol,strokecol,textcol,caption,tooltip,linkurl
' ============================================================

Public Sub BuildSvgFromCsv_PickFile()

    Dim csvPath As String

    csvPath = PickCsvFile()
    If Len(csvPath) = 0 Then Exit Sub

    BuildSvgFromCsv csvPath

End Sub

Public Sub BuildSvgFromCsv(ByVal csvPath As String)

    Dim fso As Object
    Dim svgPath As String
    Dim lines As Collection
    Dim headerFields As Variant
    Dim keyToIndex As Object
    Dim shapes As Collection
    Dim routes As Collection
    Dim i As Long
    
    Dim pageSettings As Object
    Dim rowType As String
    Dim rowData As Object
    
    
    Set pageSettings = CreateObject("Scripting.Dictionary")
    pageSettings.CompareMode = vbTextCompare


    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(csvPath) Then
        Err.Raise vbObjectError + 1000, "BuildSvgFromCsv", "CSV file not found: " & csvPath
    End If

    svgPath = fso.BuildPath( _
                fso.GetParentFolderName(csvPath), _
                fso.GetBaseName(csvPath) & ".svg" _
              )

    Set lines = ReadTextLines(csvPath)

    If lines.Count < 2 Then
        Err.Raise vbObjectError + 1001, "BuildSvgFromCsv", "CSV must contain a header row and at least one data row."
    End If

    headerFields = ParseCsvLine(CStr(lines(1)))
    Set keyToIndex = BuildHeaderMap(headerFields)

    ValidateRequiredKeys keyToIndex

    Set shapes = New Collection
    Set routes = New Collection


    For i = 2 To lines.Count
    
        If Len(Trim$(CStr(lines(i)))) = 0 Then
            GoTo ContinueLoop
        End If
    
        Set rowData = ParseShapeRow(CStr(lines(i)), keyToIndex)
    
        rowType = CStr(rowData(KEY_TYPE))
    
        If UCase$(rowType) = "_META" Then
    
            pageSettings(KEY_PAGEWIDTH) = rowData(KEY_PAGEWIDTH)
            pageSettings(KEY_PAGEHEIGHT) = rowData(KEY_PAGEHEIGHT)
            pageSettings(KEY_PAGETITLE) = rowData(KEY_PAGETITLE)
            pageSettings(KEY_BACKGROUND) = rowData(KEY_BACKGROUND)
            pageSettings(KEY_TITLECOLOUR) = rowData(KEY_TITLECOLOUR)

            pageSettings(KEY_HGRID) = rowData(KEY_HGRID)
            pageSettings(KEY_VGRID) = rowData(KEY_VGRID)
    
        ElseIf LCase$(rowType) = "rect" Then
    
            shapes.Add rowData

        ElseIf LCase$(rowType) = "route" Then
            
            routes.Add rowData

        Else
            Debug.Print "Unknown row type: " & rowType
    
        End If

ContinueLoop:

    Next i

    If shapes.Count = 0 Then
        Err.Raise vbObjectError + 1002, "BuildSvgFromCsv", "No shape rows found in CSV."
    End If

    ApplyGridLayout shapes, pageSettings

    DebugPrintRoutesFromCsv _
        shapes, _
        pageSettings, _
        routes
        
    AnalyseTextLayout shapes

    WriteSvgFile svgPath, shapes, pageSettings
    MsgBox "SVG written:" & vbCrLf & svgPath, vbInformation, "CSV-to-SVG complete"

End Sub


