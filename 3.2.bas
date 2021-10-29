Sub quadraticequation()

    Dim Document
    Dim selection
    Dim splited
    
    Dim a, b, c, d, disc, x1, x2 As Integer
    
    
    Set Document = Application.ActiveDocument
    Set selection = Document.Content
    
    Dim stringOne As String
    Dim regexOne As Object
     
    Set regexOne = New RegExp
     
    regexOne.Pattern = "([+|-]?\d+)\*x2+([+|-]?\d+)\*x([+|-]?\d+)=([+|-]?\d+)"
    regexOne.Global = True
    regexOne.IgnoreCase = IgnoreCase
    stringOne = selection

    Set theMatches = regexOne.Execute(stringOne)(0)
    
    a = CInt(theMatches.SubMatches(0))
    b = CInt(theMatches.SubMatches(1))
    c = CInt(theMatches.SubMatches(2))
    d = CInt(theMatches.SubMatches(3))
    
    Debug.Print a
    Debug.Print b
    Debug.Print c
    Debug.Print d
    
    c = c - d
    disc = b ^ 2 - 4 * a * c

    selection.Font.name = "Arial"
    selection.Font.Size = 16
    selection.Font.Italic = True

    If disc = 0 Then
        x1 = -b / (2 * a)
        selection.InsertAfter text:=vbNewLine & Replace("x = %1", "%1", CStr(x1))
    ElseIf disc > 0 Then
        x1 = (-b + disc ^ (1 / 2)) / (2 * a)
        x2 = (-b - disc ^ (1 / 2)) / (2 * a)
        selection.InsertAfter text:=vbNewLine & Replace(Replace("x1 = %1, x2 = %2", "%1", CStr(x1)), "%2", CStr(x2))
    Else
        selection.InsertAfter text:=vbNewLine & "Данное уравнение не имеет действительных чисел."
    End If
End Sub
