Sub factorial()

    Dim doc
    Dim sel

    Dim name As String
    name = VBA.InputBox("Factorial")

    Dim result As Double
    result = 1
    Dim counter As Integer

    For counter = 2 To CInt(name)
        result = result * counter
    Next counter

    Dim objSelection


    Set objSelection = Application.ActiveDocument.Content

    objSelection.Font.Size = 16
    objSelection.Font.name = "Arial"
    objSelection.Font.Italic = True
    objSelection.InsertAfter text:=CStr(result)

End Sub
