Sub ConvertTextToNumber()
    With Range("B:B")
        .NumberFormat = "General"
        .Value = .Value
    End With
End Sub
