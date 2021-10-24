'ДЛЯ КОРРЕКТНОЙ РАБОТЫ:
'объеденить клетки в диапазоне от А1 до С1 (Главная - Выравнивание - Объеденить и поместить в центре)
'написать в клетке А6 - "Сброс игры"
'написать в клетке F1 - "Имя первого игрока"
'написать в клетке G1 - "Имя второго игрока"
'выделить цветом игровое поле в диапазоне от А2 до С4

'состояние игры
'0 - игра остановлена, 1 - ход первого игрока, 2 - ход второго игрока
Private state

'выйгравший игрок
'0 - ни один из игроков (еще) не выйграл, 1 - выйграл первый игрок, 2 - выйграл второй игрок
Private win

'имя первого игрока, берется из ячейки F2
Private playerone

'имя второго игрока, берется из ячейки G2
Private playertwo

'флаг для заполненного поля
'0 - поле не заполнено, 1 - поле заполнено
Private filledflag

'при выборе новой ячейки:
Sub Worksheet_SelectionChange(ByVal Target As Range)

    'если выбрана ячейка A1
    If Not Intersect(Target, Range("A1")) Is Nothing Then
        'если игра остановлена
        If state = 0 Then
            'случайно выбрать первого игрока, который будет ходить
            state = Int((2 * Rnd) + 1)

            'оповестить о начале игры
            MsgBox "Игра началась!"

            'записать в ячейку A1 кто ходит первым
            If state = 1 Then
                Range("A1").Cells(1, 1).Value = "Ход игрока " & playerone & " (O)"
            ElseIf state = 2 Then
                Range("A1").Cells(1, 1).Value = "Ход игрока " & playertwo & " (X)"
            End If

            'записать в переменную playerone имя первого игрока из ячейки F2
            'если ячейка пустая - имя первого игрока будет "1"
            If Not Range("F2") = "" Then
                playerone = Range("F2").Value
            Else
                playerone = "1"
            End If

            'записать в переменную playertwo имя второго игрока из ячейки G2
            'если ячейка пустая - имя второго игрока будет "2"
            If Not Range("G2") = "" Then
                playertwo = Range("G2").Value
            Else
                playertwo = "2"
            End If
        End If
    End If

    'если выбрана ячейка A6 сбросить игру
    If Not Intersect(Target, Range("A6")) Is Nothing Then
        state = 0
        win = 0
        filledflag = 0
        Range("A1").Cells(1, 1).Value = "Нажмите чтобы играть"
        Call clear
        MsgBox "Игра сброшена"
    End If

    'если выбрана одна из ячеек на игровом поле в диапазоне А2-С4
    If Not Intersect(Target, Range("A2:C4")) Is Nothing Then
        'если ячейка пустая поставить в ячейку символ игрока, чей ход сейчас активен (Х или О)
        If Target.Value = "" Then
            'если ход игрока 1
            If state = 1 Then
                'поставить О в ячейку
                Target.Value = "O"
                'дать ход другому игроку
                state = 2
                Range("A1").Cells(1, 1).Value = "Ход игрока " & playertwo & " (X)"

            'если ход игрока 2
            ElseIf state = 2 Then
                'поставить Х в ячейку
                Target.Value = "X"
                'дать ход другому игроку
                state = 1
                Range("A1").Cells(1, 1).Value = "Ход игрока " & playerone & " (O)"
            End If
        End If
        
        'проверить выйграл ли один из игроков и записать в переменную win
        win = CheckIfWin()

        'если выйграл игрок 1
        If win = 1 Then
            'очистить поле
            Call clear
            
            'сбросить переменные
            state = 0
            win = 0

            'оповестить о том, что игрок 1 выйграл
            Range("A1").Cells(1, 1).Value = "Игрок " & playerone & " выйграл!"

        'если выйграл игрок 2
        ElseIf win = 2 Then

            'очистить поле
            Call clear

            'сбросить переменные
            state = 0
            win = 0

            'оповестить о том, что игрок 2 выйграл
            Range("A1").Cells(1, 1).Value = "Игрок " & playertwo & " выйграл!"

        'если не выйграл никто но поле уже заполнено - объявить ничью
        Else
            
            'проверить заполнено ли поле
            filledflag = 1
            For Each cell In Range("A2:C4")
                If cell.Value = "" Then
                    filledflag = 0
                    Exit For
                End If
            Next

            'если поле заполнено, но никто не выйграл
            If filledflag = 1 Then
                
                'сбросить переменные
                filledflag = 0
                state = 0
                win = 0

                'объявить ничью
                Range("A1").Cells(1, 1).Value = "Ничья!"

                'очистить поле
                Call clear
            End If
        End If
    End If
End Sub

'проверить выйграл ли один из игроков и вернуть соответствующее значение
Function CheckIfWin()
    '~~~~~~~~~~~~~~~~~~~~~ДИАГОНАЛИ~~~~~~~~~~~~~~~~~~~~~~~
    'основная (из левого верхнего до правого нижнего угла)
    'X O O
    'O X O
    'O O X
    If Range("A2:C4").Cells(1, 1) = Range("A2:C4").Cells(2, 2) And Range("A2:C4").Cells(2, 2) = Range("A2:C4").Cells(3, 3) And Not Range("A2:C4").Cells(1, 1) = "" And Not Range("A2:C4").Cells(2, 2) = "" And Not Range("A2:C4").Cells(3, 3) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 1))

    'побочная (из правого верхнего до левого нижнего угла)
    'O O X
    'O X O
    'X O O
    ElseIf Range("A2:C4").Cells(1, 3) = Range("A2:C4").Cells(2, 2) And Range("A2:C4").Cells(2, 2) = Range("A2:C4").Cells(3, 1) And Not Range("A2:C4").Cells(1, 3) = "" And Not Range("A2:C4").Cells(2, 2) = "" And Not Range("A2:C4").Cells(3, 1) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 3))
    
    '~~~~~~~~~~~~~~~~~~~~~~ГОРИЗОНТАЛИ~~~~~~~~~~~~~~~~~~~~~~
    'первая строка
    'X X X
    'O O O
    'O O O
    ElseIf Range("A2:C4").Cells(1, 1) = Range("A2:C4").Cells(1, 2) And Range("A2:C4").Cells(1, 2) = Range("A2:C4").Cells(1, 3) And Not Range("A2:C4").Cells(1, 1) = "" And Not Range("A2:C4").Cells(1, 2) = "" And Not Range("A2:C4").Cells(1, 3) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 1))

    'вторая строка
    'O O O
    'X X X
    'O O O
    ElseIf Range("A2:C4").Cells(2, 1) = Range("A2:C4").Cells(2, 2) And Range("A2:C4").Cells(2, 2) = Range("A2:C4").Cells(2, 3) And Not Range("A2:C4").Cells(2, 1) = "" And Not Range("A2:C4").Cells(2, 2) = "" And Not Range("A2:C4").Cells(2, 3) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(2, 1))
    
    'третья строка
    'O O O
    'O O O
    'X X X
    ElseIf Range("A2:C4").Cells(3, 1) = Range("A2:C4").Cells(3, 2) And Range("A2:C4").Cells(3, 2) = Range("A2:C4").Cells(3, 3) And Not Range("A2:C4").Cells(3, 1) = "" And Not Range("A2:C4").Cells(3, 2) = "" And Not Range("A2:C4").Cells(3, 3) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(3, 1))
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~ВЕРТИКАЛИ~~~~~~~~~~~~~~~~~~~~~
    'первый столбец
    'X O O
    'X O O
    'X O O
    ElseIf Range("A2:C4").Cells(1, 1) = Range("A2:C4").Cells(2, 1) And Range("A2:C4").Cells(2, 1) = Range("A2:C4").Cells(3, 1) And Not Range("A2:C4").Cells(1, 1) = "" And Not Range("A2:C4").Cells(2, 1) = "" And Not Range("A2:C4").Cells(3, 1) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 1))

    'второй столбец
    'O X O
    'O X O
    'O X O
    ElseIf Range("A2:C4").Cells(1, 2) = Range("A2:C4").Cells(2, 2) And Range("A2:C4").Cells(2, 2) = Range("A2:C4").Cells(3, 2) And Not Range("A2:C4").Cells(1, 2) = "" And Not Range("A2:C4").Cells(2, 2) = "" And Not Range("A2:C4").Cells(3, 2) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 2))

    'третий столбец
    'O O X
    'O O X
    'O O X
    ElseIf Range("A2:C4").Cells(1, 3) = Range("A2:C4").Cells(2, 3) And Range("A2:C4").Cells(2, 3) = Range("A2:C4").Cells(3, 3) And Not Range("A2:C4").Cells(1, 3) = "" And Not Range("A2:C4").Cells(2, 3) = "" And Not Range("A2:C4").Cells(3, 3) = "" Then
    CheckIfWin = CheckSide(Range("A2:C4").Cells(1, 3))
    
    'если никто не выйграл вернуть 0
    Else
    CheckIfWin = 0
    End If
End Function

'проверить какой из игроков выйграл
Function CheckSide(cell)
    'если Х - выйграл второй игрок (вернуть 2)
    If cell = "X" Then
        CheckSide = 2
    'если О - выйграл первый игрок (вернуть 1)
    Else
        CheckSide = 1
    End If
End Function

'очистить поле для игры
Sub clear()
    'выбрать каждую клетку в диапазоне A2-C4
    For Each cell In Range("A2:C4")
        'очистить клетку
        cell.Value = ""
    Next
End Sub
