Attribute VB_Name = "Module02"
'======================================================================================================================================================
'================================МОДУЛЬ ОПРЕДЕЛЕНИЯ ОБЩИХ СЛУЖЕБНЫХ ПОДПРОГРАММ ОТПРАВКИ SMS-СООБЩЕНИЙ И СЕРВИСА=======================================
'======================================================================================================================================================


'=================================================================ОТПРАВКА SMS=========================================================================

Function ОтправитьSMS(Сообщение, ТелНомер, ФИО, Примечание) As String  'вызов функции отправки SMS с получением Ответа выполнения операции и записью в журнал
    On Error GoTo er
    
    КодированноеСообщение = URLEncode(Сообщение): If InStr(КодированноеСообщение, "ERROR:") > 0 Then ОтправитьSMS = КодированноеСообщение: Exit Function
    
    URL_SendSMS_ = Replace(URL_SendSMS, "{login}", Логин):       URL_SendSMS_ = Replace(URL_SendSMS_, "{password}", Пароль)
    URL_SendSMS_ = Replace(URL_SendSMS_, "{phone}", ТелНомер):   URL_SendSMS_ = Replace(URL_SendSMS_, "{text}", КодированноеСообщение)

    РезультатОтправкиSMS = Trim(GetHTTPResponse(URL_SendSMS_ & IIf(РежимОтладки, "&debug=true", "")))
    
    Ошибка = ПроверкаОшибкиЗапроса(РезультатОтправкиSMS): If Ошибка <> "ok" Then ДетализацияСтатуса = Ошибка: ОтправитьSMS = Ошибка Else ДетализацияСтатуса = Replace(РезультатОтправкиSMS, "&", Chr(10))
    
    Статус = ДружелюбныйОтвет(ДетализацияСтатуса)
    
    If Ошибка = "ok" Then If КодОтвета(РезультатОтправкиSMS, 100) Then ID = Trim(Split(Split(РезультатОтправкиSMS, "&")(1), ":")(1)) Else ID = "Не верный код ответа (" & Left(Trim(Ответ), 3) & ")"
    
    Ошибка = ЗаписатьВЖурналРассылки(Сообщение, ТелНомер, ФИО, Примечание, Статус, ДетализацияСтатуса, ID) 'Unicode
    
    If Ошибка <> "" Then MsgBox Ошибка & p(2) & "При попытке отправки SMS сервер вернул сообщение: " & Статус

    If ОтправитьSMS = "" Then ОтправитьSMS = "ok"
    
    Exit Function
er:
    ОтправитьSMS = "Ошибка отправки SMS: " & Err.Source & ": " & Err.Number & ": " & Err.Description & ": " & "URL: " & URL_SendSMS_
End Function


    Function ЗаписатьВЖурналРассылки(Сообщение, ТелНомер, ФИО, Примечание, Статус, ДетализацияСтатуса, ID)
        On Error GoTo er
        
        НижняяЗаполненнаяСтрока = НижняяЗаполненнаяЯчейка(Sheets("Журнал рассылки")): If НижняяЗаполненнаяСтрока < 5 Then Exit Function
        
        pr "Журнал рассылки", 0
        
        With Sheets("Журнал рассылки")
            With .Range("a6").Offset(НижняяЗаполненнаяСтрока - 5, 0)
                With .Range(.Row & ":" & .Row): .Font.Color = RGB(0, 0, 0): .Interior.TintAndShade = 0: End With
                
                .value = Format(Date, "dd.mm.yyyy"): .Offset(0, 1).value = Format(Time, "hh:mm:ss"): .Offset(0, 2).value = ТелНомер: .Offset(0, 3).value = Сообщение
                .Offset(0, 4).value = ФИО: .Offset(0, 5).value = Примечание: .Offset(0, 6).value = Статус: .Offset(0, 7).value = ДетализацияСтатуса: .Offset(0, 8).value = ID
            End With
        End With
        
        pr "Журнал рассылки"
        
        Exit Function
er:
        ЗаписатьВЖурналРассылки = "Записать действие в журнал рассылки не удалось! Ошибка: " & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description
        
        pr "Журнал рассылки"
    End Function


Function ОбщаяПодготовкаРассылки()
    On Error GoTo er
    
    If Not ЗаполнениеМассивовСпискаРассылки Then Exit Function Else If Not ПроверкаАвторизации Then Exit Function
    
    ОбщаяПодготовкаРассылки = True
    
    Exit Function
er:
    MsgBox "Ошибка общей подготовки рассылки!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ПроверкаАвторизации()
    On Error GoTo er
    
    If Пароль = "" Or Логин = "" Then ФормаВводаПароля.Show
    If Пароль = "" Or Логин = "" Then MsgBox "Номер телефона и пароль не могут быть пустыми!", vbExclamation: Exit Function
    
    ПроверкаАвторизации = True
    
    Exit Function
er:
    MsgBox "Ошибка проверки авторизации!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ЗаполнениеМассивовСпискаРассылки()
    On Error GoTo er
    
    Dim reg As Object, math As Object, i
    
    НижняяЗаполненнаяСтрока = НижняяЗаполненнаяЯчейка(Sheets("Список рассылки")): If НижняяЗаполненнаяСтрока < 5 Then Exit Function
        
    ОшибкаНомера = False: Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^79\d{9}$" 'маска проверки номера
    
    With Sheets("Список рассылки")
        ReDim ТелНомера(0): ReDim ФлагиАктивностиТелНомера(0): ReDim ФИОы(0): ReDim Примечания(0)
            
        For i = 6 To НижняяЗаполненнаяСтрока
            With .Range("a1").Offset(i - 1, 0)
                ТелНомер = Trim(.value): ФлагАктивностиТелНомера = .Font.Bold
                
                If ТелНомер <> "" And ФлагАктивностиТелНомера Then
                    Set math = reg.Execute(ТелНомер)
                    
                    If math.Count = 1 Then
                        .Font.Color = RGB(0, 0, 0)
                        
                        If ТелНомера(LBound(ТелНомера)) <> "" Then ReDim Preserve ТелНомера(UBound(ТелНомера) + 1): ReDim Preserve ФИОы(UBound(ФИОы) + 1): ReDim Preserve Примечания(UBound(Примечания) + 1):
                    
                        ТелНомера(UBound(ТелНомера)) = ТелНомер: ФИОы(UBound(ФИОы)) = Trim(.Offset(0, 1).value): Примечания(UBound(Примечания)) = Trim(.Offset(0, 2).value)
                    Else
                        .Font.Color = RGB(255, 0, 0): ОшибкаНомера = True
                    End If
                End If
            End With
        Next
    End With
    
    If ОшибкаНомера Then: MsgBox "Внимание! В списке рассылки присутствуют номера не правильного формата. Формат номера должен быть таким: 79998887766", vbExclamation: Sheets("Список рассылки").Activate
    
    If ТелНомера(LBound(ТелНомера)) = "" Then MsgBox "В списке рассылки ни одного правильного или активного номера!", vbExclamation: Sheets("Список рассылки").Activate: Exit Function
    
    Set reg = Nothing: Set math = Nothing
    
    ЗаполнениеМассивовСпискаРассылки = True
    
    Exit Function
er:
    MsgBox "Ошибка обработки номеров списка расслки" & IIf(i > 0, " в строке " & i, "") & "!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ПроверкаДатыСтатусаДоставки()
    On Error GoTo er
    
    Dim reg As Object, math As Object, i
    
    Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^\d{2}\.\d{2}\.\d{4}$" 'маска проверки даты
    
    Set math = reg.Execute(Trim(Sheets("Журнал рассылки").Range("c3"))):  If math.Count <> 1 Then MsgBox "Введенная дата статуса доставки должна быть в формате: ДД.ММ.ГГГГ, например: " & Format(Date, "dd.mm.yyyy"), vbExclamation: Exit Function
    
    If Not IsDate(math(0)) Then MsgBox "Введенная дата статуса не верна", vbCritical: Exit Function

    ПроверкаДатыСтатусаДоставки = True
    
    Exit Function
er:
    MsgBox "Ошибка проверки даты статуса доставки" & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Sub ПроверкаСтатусаДоставки()
	If ПроверкаНаСуществованиеВсехКомпонентовКниги (1) <> "::" Then Exit Sub

    ПроверкаСтатусаДоставки_
End Sub


Sub ПроверкаСтатусаДоставки_()
    On Error GoTo er
    
    If Not ПроверкаДатыСтатусаДоставки Then Exit Sub Else If Not ПроверкаАвторизации Then Exit Sub
    
    If Not ЧтениеНастроек Then Exit Sub Else If Not ПроверкаГотовности Then Exit Sub
    
    НижняяЗаполненнаяСтрока = НижняяЗаполненнаяЯчейка(Sheets("Журнал рассылки")): If НижняяЗаполненнаяСтрока < 5 Then Exit Sub
        
    КрасныйШрифт = RGB(255, 0, 0): ЧерныйШрифт = RGB(0, 0, 0)
    
    КраснаяЗаливка = RGB(255, 200, 200): ЖелтаяЗаливка = RGB(255, 255, 200): ЗеленаяЗаливка = RGB(200, 255, 200)
        
    pr "Журнал рассылки", 0
    
    With Sheets("Журнал рассылки")
        Дата = .Range("c3").value: Счетчик = 0
    
        URL_StatusSMS_ = Replace(URL_StatusSMS, "{login}", Логин):  URL_StatusSMS_ = Replace(URL_StatusSMS_, "{password}", Пароль)
    
        For Each Ячейка In .Range("a6:a" & НижняяЗаполненнаяСтрока)
            ОшибкаВСтроке = False: ЦветЗаливки = 0: ПрежнееЗначениеСчетчика = Счетчик
            
            If IsDate(Ячейка.value) Then
                ID = .Range("i" & Ячейка.Row).value
                    
                If ID <> "" And CDate(Ячейка.value) = CDate(Дата) And Not (Ячейка.Interior.Color = КраснаяЗаливка Or Ячейка.Interior.Color = ЗеленаяЗаливка) Then
                    If IsNumeric(ID) Then
                        Ответ = Trim(GetHTTPResponse(Replace(URL_StatusSMS_, "{id}", ID))): Ошибка = ПроверкаОшибкиЗапроса(Ответ)
                        
                        If Ошибка = "ok" And КодОтвета(Ответ, 101) Then
                            Статус = СтатусДоставки(Ответ): If Статус <> "" Then Счетчик = Счетчик + 1
                            
                            Select Case Статус
                                Case "доставлено": ЦветЗаливки = ЗеленаяЗаливка: Case "донедоставлено": ЦветЗаливки = ЖелтаяЗаливка: Case "недоставлено": ЦветЗаливки = КраснаяЗаливка
                            End Select
                         Else: ОшибкаВСтроке = True: End If
                    Else: ОшибкаВСтроке = True: End If
                End If
            Else: ОшибкаВСтроке = True: End If
            
            With .Range(Ячейка.Row & ":" & Ячейка.Row)
                If Not ОшибкаВСтроке Then
                    .Font.Color = ЧерныйШрифт: If ЦветЗаливки <> 0 And ПрежнееЗначениеСчетчика < Счетчик Then .Interior.Color = ЦветЗаливки Else .Interior.TintAndShade = 0
                Else
                    .Font.Color = КрасныйШрифт: .Interior.TintAndShade = 0
                End If
            End With
        Next
        
        Сообщение = "Для выбранной даты SMS-сообщений возвращен статус доставки/недоставки для " & Счетчик & " строк."
        
        If Счетчик = 0 Then MsgBox Сообщение & p(2) & "Возможные причины:" & p(1) & "    1) статус SMS-сообщений еще не известен;" & p(1) & _
                                   "    2) сообщений без статусов на выбранную дату не обнаружено;" & p(1) & "    3) в строках журнала имеются " & _
                                   "ошибки;" & p(1) & "    4) сервер не вернул идентификаторы SMS-сообщений.", vbExclamation _
                       Else _
                            MsgBox Сообщение, vbInformation
    End With
    
    pr "Журнал рассылки"
    
    Exit Sub
er:
    MsgBox "Ошибка проверки статуса доставки" & ": " & Err.Number & ": " & Err.Description, vbCritical
    
    pr "Журнал рассылки"
End Sub

    
    Function СтатусДоставки(Ответ)
        If InStr(Ответ, ":") > 0 Then
            Статус = Trim(Split(Split(Ответ, "&")(1), ":")(0)): Дата = Trim(Split(Split(Ответ, "&")(1), ":")(1))
        Else
            Статус = Trim(Split(Ответ, "&")(1))
        End If
        
        If LCase(Статус) = "доставлено" Then
            If Дата <> "" Then СтатусДоставки = "доставлено" Else СтатусДоставки = "донедоставлено"
            
        ElseIf LCase(Статус) = "недоставлено" Then
            СтатусДоставки = "недоставлено"
        End If
    End Function


Function ПроверкаЗапрещенныхВхождений(Строка)
    On Error Resume Next
    
    ЗапрещенныеВхождения = "": If Строка = "" Then Exit Function

    МассивЗапрещенныхВхождений = Split(СтрокаЗапрещенныхВхожденийВТекстСообщения, " ")

    For i = LBound(МассивЗапрещенныхВхождений) To UBound(МассивЗапрещенныхВхождений)
        If InStr(Строка, МассивЗапрещенныхВхождений(i)) > 0 Then ЗапрещенныеВхождения = ЗапрещенныеВхождения & " " & МассивЗапрещенныхВхождений(i)
    Next
    
    For x = 1 To Len(Строка)
        If Симв(Mid(Строка, x, 1)) = "" Then ЗапрещенныеВхождения = ЗапрещенныеВхождения & " " & Mid(Строка, x, 1)
    Next
    
    If ЗапрещенныеВхождения <> "" Then MsgBox "В тексте сообщения запрещено использовать: " & ЗапрещенныеВхождения, vbExclamation: Exit Function
    
    ПроверкаЗапрещенныхВхождений = True
End Function


Function ПроверкаОшибкиЗапроса(Ответ)
    On Error Resume Next
    
    ПроверкаОшибкиЗапроса = "ok"
    
    If Trim(Ответ) = "" Then ПроверкаОшибкиЗапроса = "Web-сервис вернул пустой ответ"
    
    If InStr(Ответ, "ERROR: GetHTTPResponse завершилась с ошибкой:") > 0 Then 'ошибка функции GetHTTPResponse
        ПроверкаОшибкиЗапроса = "Внутренняя ошибка механизма запроса к вэб-серверу. " & _
                                "Проверьте соединение с интернетом." & p2(2) & Ответ
        
    ElseIf InStr(Ответ, "TIMEOUT") > 0 Then 'сервис не доступен
        ПроверкаОшибкиЗапроса = "Вэб-сервис временно недоступен. Попробуйте позже." & p2(2) & Ответ
        
    ElseIf InStr(Ответ, "ERROR:") > 0 Or КодОтвета(Ответ, 400, ">=") Then 'ошибка сервиса уровня сервера
        ПроверкаОшибкиЗапроса = "Ошибка веб-сервиса. Обратитесь к администратору." & p2(2) & Ответ
        
    ElseIf КодОтвета(Ответ, 300, ">=") Then 'ошибка сервиса уровня клиента
        ПроверкаОшибкиЗапроса = Ответ
    
        If КодОтвета(Ответ, 301, ">=") And КодОтвета(Ответ, 304, "<=") Then Логин = "": Пароль = ""
    End If
End Function


Function КодОтвета(ByVal Ответ, Код As Integer, Optional Сравнение = "=") As Boolean
    On Error Resume Next
    
    'ищет первое совпадение сравнения с Код'ом
    
    КодОтвета = False: Ответ = Trim(Ответ): If Ответ = "" Then Exit Function
    
    СообщенияСервиса = Split(Ответ, "&")
    
    For i = LBound(Сообщения) To UBound(Сообщения)
        СообщениеСервиса = Trim(СообщенияСервиса(i)): Члены = Split(СообщениеСервиса, ":")
        
        If UBound(Члены) - LBound(Члены) + 1 = 2 Then
            КодСообщения = Trim(Члены(LBound(Члены)))
            
            If IsNumeric(КодСообщения) Then
                Select Case Сравнение
                    Case "=":  КодОтвета = (CInt(КодСообщения) = Код):  Case ">=": КодОтвета = (CInt(КодСообщения) >= Код):  Case "<=": КодОтвета = (CInt(КодСообщения) <= Код)
                    Case ">":  КодОтвета = (CInt(КодСообщения) > Код):  Case "<":  КодОтвета = (CInt(КодСообщения) < Код)
                End Select
            End If
        End If
        
        If КодОтвета Then Exit For
    Next
End Function


Function ДружелюбныйОтвет(Ответ)
    On Error GoTo er
    
    Select Case Left(Trim(Ответ), 3)
        Case "100": ДружелюбныйОтвет = Split(Split(Ответ, Chr(10))(0), ":")(1) & " стоимостью " & ВычислениеРублейИКопеек(Split(Split(Ответ, Chr(10))(2), ":")(1))
        
        Case Else: If IsNumeric(Left(Trim(Ответ), 3)) And InStr(Ответ, ":") > 0 Then ДружелюбныйОтвет = Split(Ответ, ":")(1) Else ДружелюбныйОтвет = Ответ
    End Select
    
    Exit Function
er:
    ДружелюбныйОтвет = Ответ
End Function


Sub ЗаписатьВЛог(Текст, Optional Ответ = False)
    On Error Resume Next
    
    With Sheets("Лог")
        НижняяЗаполненнаяСтрока = НижняяЗаполненнаяЯчейка(Sheets("Лог")): If НижняяЗаполненнаяСтрока < 2 Then Exit Sub
        
        НижняяЗаполненнаяСтрока = НижняяЗаполненнаяСтрока - IIf(Ответ, 1, 0)
        
        With .Range("a1").Offset(НижняяЗаполненнаяСтрока, 0)
            If Ответ Then Сдвиг = 3
            
            .Offset(0, 0 + Сдвиг).value = Format(Date, "dd.mm.yyyy"): .Offset(0, 1 + Сдвиг).value = Format(Time, "hh:mm:ss"): .Offset(0, 2 + Сдвиг).value = Текст
        End With
    End With
End Sub



'====================================================================СЕРВИС============================================================================

Sub ЗапросБаланса() 'запрос баланса
    If ПроверкаНаСуществованиеВсехКомпонентовКниги (1) <> "::" Then Exit Sub

    ЗапросБаланса_
End Sub

    
Sub ЗапросБаланса_() 'запрос баланса
    On Error GoTo er
    
    Sheets("Сервис").Range("d3").value = "": If Not ЧтениеНастроек Then Exit Sub

    If Not ПроверкаАвторизации Then Exit Sub Else If Not ПроверкаГотовности Then Exit Sub
    
    URL_Balance_ = Replace(URL_Balance, "{login}", Логин): URL_Balance_ = Replace(URL_Balance_, "{password}", Пароль)
    
    РезультатЗапросаБаланса = Trim(GetHTTPResponse(URL_Balance_))
    
    Ошибка = ПроверкаОшибкиЗапроса(РезультатЗапросаБаланса): If Ошибка <> "ok" Then MsgBox "Произошла ошибка:" & p(2) & Ошибка: Exit Sub
    
    Sheets("Сервис").Range("d3").value = ВычислениеРублейИКопеек(Split(Split(РезультатЗапросаБаланса, "&")(1), ":")(1))
    
    Exit Sub
er:
    MsgBox "Ошибка запроса баланса!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description & p(2) & "URL: " & URL_SendSMS_, vbCritical
End Sub

    
    Function ВычислениеРублейИКопеек(Число)
        On Error GoTo er
    
        Число = Replace(Число, ".", ",")
        
        Рубли = "": If InStr(Число, ",") > 0 Then Рубли = Split(Число, ",")(0) Else Рубли = Число
        
        Копейки = "": If InStr(Число, ",") > 0 Then Копейки = Split(Число, ",")(1): If Len(Копейки) = 1 Then Копейки = Копейки & "0"
        
        ВычислениеРублейИКопеек = IIf(Рубли <> "", Рубли & " рублей ", "") & IIf(Копейки <> "", Копейки & " копеек", "")
    
        Exit Function
er:
        MsgBox "Ошибка вычисления выражениея денежных средств!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
    End Function





