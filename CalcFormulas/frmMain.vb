Imports System.Text.RegularExpressions
Imports System.Reflection

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        MsgBox(New Formula(Me.txtFormula.Text).Value)
    End Sub


End Class


Public Class Formula
    Dim _value As Decimal = Nothing
    Dim _formula As String = ""
    Dim _error As String = ""

    Public Sub New(ByVal formula As String)
        Me.GetFormulaResult(formula)
    End Sub
    Public ReadOnly Property Value As Decimal
        Get
            Return _value
        End Get
    End Property
    Public Property Formula As String
        Get
            Return _formula
        End Get
        Set(value As String)
            _formula = value
            GetFormulaResult(value)
        End Set
    End Property
    Public ReadOnly Property [Error] As String
        Get
            Return _error
        End Get
    End Property

    Public Function GetFormulaResult(ByVal formula As String) As Decimal
        _error = ""
        _formula = formula

        Dim opz As String = ПреобразоватьФормулуВОПЗ(formula)
        _value = ПолучитьРезультатВычисленияИзОПЗ(opz)

        If Not String.IsNullOrEmpty(_error) Then
            MsgBox(_error)
            Return Nothing
        End If

        Return _value
    End Function

    Private Function ПолучитьПриоритетОперации(ByVal oper As Char) As Integer
        If oper = "^" Then
            Return 2
        ElseIf oper = "*" Or oper = "/" Then
            Return 1
        ElseIf oper = "+" Or oper = "-" Then
            Return 0
        Else : Return -1
        End If
    End Function
    Private Function ПреобразоватьФормулуВОПЗ(ByVal formula As String) As String
        formula = formula.Replace(" ", "").Replace(",", ".") 'Убираем из выражения пробелы, и заменяем все запятые на точки - разделители десятичных

        ' Перед преобразованием формулы в ОПЗ необходимо рассчитать значения всех функций используемых в ней
        formula = РассчитатьЗначенияФункций(formula)
        If Not String.IsNullOrEmpty(_error) Then Return ""

        Dim strExit() As String = {} 'Выход
        Dim strStack() As String = {} 'Стэк


        Dim prevSymbolIsDigit As Boolean = False

        ' Начинаем посимвольно перебирать строку
        For i As Integer = 0 To formula.Length - 1
            Dim symbol As Char = formula(i)
            Dim prev_symbol As Char = Nothing
            If i > 0 Then prev_symbol = formula(i - 1)

            If Char.IsDigit(symbol) Or _
                symbol = "," Or _
                symbol = "." Or _
                symbol = "-" AndAlso ( _
                    prev_symbol = Nothing OrElse _
                    "+-*/^(".Contains(prev_symbol)) Then

                If Not prevSymbolIsDigit Then ReDim Preserve strExit(strExit.Length)

                strExit(strExit.Length - 1) += symbol
                prevSymbolIsDigit = True
            Else
                If "+-*/^()".Contains(symbol) Then ' symbol = "+" Or symbol = "-" Or symbol = "*" Or symbol = "/" Or symbol = "^" Or symbol = "(" Or symbol = ")" Then
                    If strStack.Length >= 1 And ПолучитьПриоритетОперации(symbol) >= 0 AndAlso _
                        (ПолучитьПриоритетОперации(symbol) < ПолучитьПриоритетОперации(strStack(strStack.Length - 1)) Or _
                        ПолучитьПриоритетОперации(symbol) = 1 And _
                        ПолучитьПриоритетОперации(strStack(strStack.Length - 1)) = 1) Then

                        ReDim Preserve strExit(strExit.Length)
                        strExit(strExit.Length - 1) = strStack(strStack.Length - 1)
                        strStack(strStack.Length - 1) = symbol
                    ElseIf symbol = ")" Then
                        Do
                            If strStack.Length = 0 Then MsgBox("В выражении не согласованы скобки: " + formula) : Return ""
                            If strStack(strStack.Length - 1) = "(" Then
                                ReDim Preserve strStack(strStack.Length - 2)
                                Exit Do
                            Else
                                ReDim Preserve strExit(strExit.Length)
                                strExit(strExit.Length - 1) = strStack(strStack.Length - 1)
                                ReDim Preserve strStack(strStack.Length - 2)
                            End If
                        Loop
                    Else
                        ReDim Preserve strStack(strStack.Length)
                        strStack(strStack.Length - 1) = symbol
                    End If
                End If
                prevSymbolIsDigit = False
            End If
        Next

        For i As Integer = strStack.Length - 1 To 0 Step -1
            If strStack(i) = "(" Or strStack(i) = ")" Then MsgBox("В выражении не согласованы скобки: " + formula) : Return ""
            ReDim Preserve strExit(strExit.Length)
            strExit(strExit.Length - 1) = strStack(i)
        Next

        Dim str As String = ""
        'For i As Integer = 0 To strExit.Length - 1
        '    str += strExit(i) + IIf(i < strExit.Length - 1, ":", "")
        'Next

        str = String.Join(":", strExit)

        Return str
    End Function
    Private Function ПолучитьРезультатВычисленияИзОПЗ(ByVal OPZ As String) As Decimal
        Dim strStack() As String = OPZ.Split(":")

        Dim i As Integer = 0
        Do
            If i > strStack.Length - 1 Then
                MsgBox("Ошибка строки ПОЛИЗ при расчете формулы: Индекс за пределами массива." + vbNewLine + _
                        "Вероятно в расчет расхода комплектующих был некорректно добавлен материал.")
                Return 0
            End If
            If strStack(i) = "+" Or strStack(i) = "-" Or strStack(i) = "/" Or strStack(i) = "*" Or strStack(i) = "^" Then
                If i < 2 Then MsgBox("Ошибка строки ПОЛИЗ при расчете формулы: " + OPZ) : Return 0

                Dim first As Decimal = Val(strStack(i - 2).Replace(",", "."))
                Dim second As Decimal = Val(strStack(i - 1).Replace(",", "."))

                If strStack(i) = "+" Then
                    strStack(i - 2) = (first + second).ToString
                ElseIf strStack(i) = "-" Then
                    strStack(i - 2) = (first - second).ToString
                ElseIf strStack(i) = "*" Then
                    strStack(i - 2) = (first * second).ToString
                ElseIf strStack(i) = "/" Then
                    strStack(i - 2) = (first / second).ToString
                ElseIf strStack(i) = "^" Then
                    strStack(i - 2) = (first ^ second).ToString
                End If

                For k As Integer = i + 1 To strStack.Length - 1
                    strStack(k - 2) = strStack(k)
                Next

                ReDim Preserve strStack(strStack.Length - 3)
                i -= 2
            Else
                i += 1
            End If
        Loop While strStack.Length <> 1

        Return Val(strStack(0).Replace(",", "."))
    End Function

    Private Function РассчитатьЗначенияФункций(ByVal formula As String) As String
        ' Для поиска функций воспользуемся регулярным выражением, которое позволит находить только
        ' вложенные функции, потому что рассчитывать нужно сначала их
        Do
            Dim matches As MatchCollection = Regex.Matches(formula, "\w+\([\d+-^\*\/.]+(;[\d+-^\*\/.]+)*\)")
            If matches.Count = 0 Then Exit Do

            Dim match As Match = matches(0)

            Dim funcname As String = Regex.Match(match.Value, "\w+(?=\()").Value.ToUpper ' Получаем название функции
            Dim params As String() = Regex.Match(match.Value, "(?<=\w+\().*(?=\))").Value.Split(";") ' Получаем массив параметров

            Dim funcvalue As String = РассчитатьЗначениеФункции(funcname, params)

            If Not String.IsNullOrEmpty(_error) Then Return ""

            ' Заменяем в формуле функцию на полученный результат
            formula = formula.Replace(match.Value, funcvalue)
        Loop



        Return formula
    End Function
    Private Function РассчитатьЗначениеФункции(ByVal funcname As String, ByVal params As String()) As String
        ' Для упрощения создания новых функций условимся называть функции следующим образом Функция_<ИМЯФУНКЦИИ>
        ' Теперь для начала использования достаточно добавить Public-функцию с нужным именем и начать использовать её.
        Dim result = ВызватьМетодЕслиОнЕсть(Me, "Функция_" + funcname, {params}) ' А здесь будем использовать функцию по имени
        If result = False And Not String.IsNullOrEmpty(_error) Then
            Dim strError As String = String.Format("Функция {0} не задана! ", funcname)
            Me._error += strError
            Return Nothing
        End If

        Return result

    End Function

    Private Function ВызватьМетодЕслиОнЕсть(ByRef objSource As Object, _
                                     ByRef methodName As String, _
                                     ByVal params As Object()) As Object
        Dim methInfo As MethodInfo = objSource.GetType.GetMethod(methodName)
        If methInfo IsNot Nothing Then
            Return methInfo.Invoke(objSource, params)
        End If
        Return False
    End Function

    Public Function MyFunction(ByVal params As String()) As String
        Return ""
    End Function


#Region "Пользовательские функции в формулах"

    Public Function Функция_КОРЕНЬ(ByVal params As String()) As Decimal
        ' Используется 2 параметра. Если количество параметров не совпадает - ошибка
        If params.Length < 2 Then
            _error += "Функция КОРЕНЬ - должно быть 2 параметра. "
            Return Nothing
        End If

        ' Если параметров больше - то используем первые 2 - остальные игнорируем
        Dim param1, param2 As Decimal

        param1 = РассчитатьПараметр(params(0))
        param2 = РассчитатьПараметр(params(1))

        Dim result As Decimal = Nothing
        Try
            result = Math.Pow(param1, 1 / param2)
        Catch ex As Exception
            _error += ex.ToString
        End Try
        Return result
    End Function
    Public Function Функция_ОКРУГЛ(ByVal params As String()) As Decimal
        ' Используется 2 параметра. Если количество параметров не совпадает - ошибка
        If params.Length < 1 Then
            _error += "Функция ОКРУГЛ - должен быть минимум 1 параметр. "
            Return Nothing
        End If

        Dim param1, param2 As Decimal

        param1 = РассчитатьПараметр(params(0))
        If params.Length > 1 Then param2 = РассчитатьПараметр(params(1))

        Dim result As Decimal = Nothing
        Try
            Dim koef As Decimal = 10 ^ param2
            result = Math.Round(param1 * koef) / koef
        Catch ex As Exception
            _error += ex.ToString
        End Try
        Return result
    End Function
#End Region

    Private Function РассчитатьПараметр(param As String) As Decimal
        If CStr(Val(param)) = param Then ' Ничего не надо рассчитывать - параметр не формула!
            Return Val(param)
        Else
            ' Вероятно параметр задан также формулой - рассчитываем значение
            ' Поскольку мы рассчитываем функцию максимальной вложенности, т.е. в параметрах
            ' функции быть не может, можно не переживать о зацикленности.
            Dim formula As New Formula(param)
            If String.IsNullOrEmpty(formula.Error) Then
                Return formula.Value
            End If

            Return Nothing
        End If
    End Function


End Class