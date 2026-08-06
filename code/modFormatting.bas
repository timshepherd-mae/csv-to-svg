Public Function ToDbl(ByVal valueText As String) As Double

    Dim decimalSep As String
    Dim normalised As String

    decimalSep = Application.International(xlDecimalSeparator)
    normalised = Trim$(valueText)

    If decimalSep <> "." Then
        normalised = Replace(normalised, ".", decimalSep)
    End If

    ToDbl = CDbl(normalised)

End Function

Public Function SvgNum(ByVal value As Double) As String

    Dim s As String

    s = CStr(value)

    s = Replace(s, Application.International(xlDecimalSeparator), ".")

    SvgNum = s

End Function


