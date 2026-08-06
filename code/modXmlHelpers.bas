Public Function XmlSafeText(ByVal valueText As String) As String

    Dim s As String

    s = valueText
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")

    XmlSafeText = s

End Function

Public Function XmlSafeAttribute(ByVal valueText As String) As String

    Dim s As String

    s = XmlSafeText(valueText)
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")

    XmlSafeAttribute = s

End Function


