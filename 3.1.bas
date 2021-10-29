Sub factorial()
Dim a
Dim inputnum As Integer

a = Split(VBA.InputBox("Введите число"), " ")
inputnum = getResult(a)

Dim result As Variant
result = 1
Dim counter As Integer

For counter = 2 To CInt(inputnum)
If inputnum > 170 Then
result = -1
Exit For
End If
result = result * counter
Next counter

Dim objSelection
Set objSelection = Application.ActiveDocument.Content

objSelection.Font.Size = 16
objSelection.Font.name = "Arial"
objSelection.Font.Italic = True
objSelection.InsertAfter Text:=vbNewLine & Replace("Введенное число: %1", "%1", CStr(inputnum))

If result = -1 Then
objSelection.InsertAfter Text:=vbNewLine & "Невозможно вычислить факториал"
Else
objSelection.InsertAfter Text:=vbNewLine & Replace("Факториал числа: %1", "%1", CStr(result))
End If
End Sub
Public Function getResult(a)
Dim Item
Dim result
result = 0
For Each Item In a
Debug.Print Item
result = result + CInt(ReplaceNum(CStr(Item)))
Next
getResult = result
End Function
Public Function ReplaceNum(number As String) As String
Select Case number
Case "Один"
ReplaceNum = "1"

Case "Два"
ReplaceNum = "2"

Case "Три"
ReplaceNum = "3"

Case "Четыре"
ReplaceNum = "4"

Case "Пять"
ReplaceNum = "5"

Case "Шесть"
ReplaceNum = "6"

Case "Семь"
ReplaceNum = "7"

Case "Восемь"
ReplaceNum = "8"

Case "Девять"
ReplaceNum = "9"

Case "Десять"
ReplaceNum = "10"

Case "Одиннадцать"
ReplaceNum = "11"

Case "Двенадцать"
ReplaceNum = "12"

Case "Тринадцать"
ReplaceNum = "13"

Case "Четырнадцать"
ReplaceNum = "14"

Case "Пятнадцать"
ReplaceNum = "15"

Case "Шестнадцать"
ReplaceNum = "16"

Case "Семнадцать"
ReplaceNum = "17"

Case "Восемнадцать"
ReplaceNum = "18"

Case "Девятнадцать"
ReplaceNum = "19"

Case "Двадцать"
ReplaceNum = "20"

Case "Тридцать"
ReplaceNum = "30"

Case "Сорок"
ReplaceNum = "40"

Case "Пятьдесят"
ReplaceNum = "50"

Case "Шестьдесят"
ReplaceNum = "60"

Case "Семьдесят"
ReplaceNum = "70"

Case "Восемьдесят"
ReplaceNum = "80"

Case "Девяносто"
ReplaceNum = "90"

Case "Сто"
ReplaceNum = "100"

Case "Двести"
ReplaceNum = "200"

Case "Триста"
ReplaceNum = "300"

Case "Четыресто"
ReplaceNum = "400"

Case "Пятьсот"
ReplaceNum = "500"

Case "Шестьсот"
ReplaceNum = "600"

Case "Семьсот"
ReplaceNum = "700"

Case "Восемьсот"
ReplaceNum = "800"

Case "Девятьсот"
ReplaceNum = "900"

Case "Тысяча"
ReplaceNum = "1000"
End Select
End Function
