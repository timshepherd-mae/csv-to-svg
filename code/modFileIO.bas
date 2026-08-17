Public Function PickCsvFile() As String

    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select CSV shape data"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"

        If .Show <> -1 Then
            PickCsvFile = vbNullString
        Else
            PickCsvFile = .SelectedItems(1)
        End If
    End With

End Function


Public Function ReadTextLines(ByVal filePath As String) As Collection

    Dim result As New Collection
    Dim fileNum As Integer
    Dim lineText As String

    fileNum = FreeFile

    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        result.Add lineText
    Loop

    Close #fileNum

    Set ReadTextLines = result

End Function

Public Sub WriteSvgFile( _
    ByVal svgPath As String, _
    ByVal shapes As Collection, _
    ByVal pageSettings As Object, _
    ByVal routes as Collection)

    Dim fileNum As Integer
    Dim svgText As String

    svgText = BuildSvgText(shapes, pageSettings, routes)

    ' DEBUG
    ' Debug.Print svgText

    fileNum = FreeFile

    Open svgPath For Output As #fileNum
    Print #fileNum, svgText
    Close #fileNum

End Sub
