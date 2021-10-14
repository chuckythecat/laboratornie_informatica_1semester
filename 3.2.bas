Sub quadraticequation()

    Dim Document
    Dim selection
    Dim splited
    
    Dim a, b, c, d, disc, x1, x2 As Integer
    
    
    Set Document = Application.ActiveDocument
    Set selection = Document.Content
    splited = Split(selection, "+")
    
    a = CInt(Split(splited(0), "*")(0))
    b = CInt(Split(splited(1), "*")(0))
    c = CInt(Split(splited(2), "=")(0))
    d = CInt(Split(splited(2), "=")(1))
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print d
    
    c = c - d
    disc = b ^ 2 - 4 * a * c
    
    x1 = (-b + disc ^ (1 / 2)) / (2 * a)
    x2 = (-b - disc ^ (1 / 2)) / (2 * a)
    
    selection.Font.name = "Arial"
    selection.Font.Size = 16
    selection.Font.Italic = True
    
    selection.InsertAfter text:=vbNewLine & Replace(Replace("x1 = %1, x2 = %2", "%1", CStr(x1)), "%2", CStr(x2))

End Sub
