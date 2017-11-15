Attribute VB_Name = "Module08"
'======================================================================================================================================================
'======================================МОДУЛЬ ОПРЕДЕЛЕНИЯ НИЗКОУРОВНЕВОГО МЕХАНИЗМА ОТПРАВКИ SMS-СООБЩЕНИЙ=============================================
'======================================================================================================================================================


'==========================================================ВЫПОЛНЕНИЕ HTTP-ЗАПРОСОВ====================================================================

Function GetHTTPResponse(ByVal sURL As String, Optional sParam As String = "") As String 'асинхронный вызов
    On Error GoTo er
    
    Const TIMEOUT& = 10  'в секундах
    
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    If sParam <> "" Then sParam = sParam & "&randomval=" & gen(32767) Else sURL = sURL & IIf(InStr(sURL, "?") > 0, "&", "?") & "randomval=" & gen(32767)
 
    ЗаписатьВЛог sURL & IIf(sParam <> "", "?" & sParam, "")
    
    If sParam = "" Then
        winHttpReq.Open "GET", sURL, True: DoEvents: winHttpReq.Send: DoEvents
    Else
        winHttpReq.Open "POST", sURL, True: DoEvents
        winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        winHttpReq.Send sParam: DoEvents
    End If
 
    If Not winHttpReq.WaitForResponse(TIMEOUT&) Then GetHTTPResponse = "ERROR: GetHTTPResponse: Время ожидания вышло, TIMEOUT = " & TIMEOUT&: Exit Function
    If winHttpReq.Status <> "200" Then GetHTTPResponse = "ERROR: GetHTTPResponse: " & winHttpReq.Status & ": " & winHttpReq.StatusText: Exit Function
     
    GetHTTPResponse = winHttpReq.responsetext
    
    ЗаписатьВЛог GetHTTPResponse, True
    
    Set winHttpReq = Nothing
    
    Exit Function
er:
    GetHTTPResponse = "ERROR: GetHTTPResponse завершилась с ошибкой: " & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description
    
    ЗаписатьВЛог "Ответ получить не удалось: " & GetHTTPResponse, True
End Function


'============================================================КОДИРОВАНИЕ URLENCODE======================================================================

Public Function URLEncode(ByVal Текст As String) 'преобразование сообщения в кодированную строку для возможности передачи любых символов
    On Error Resume Next
    
    For x = 1 To Len(Текст): URLEncode = URLEncode & Симв(Mid(Текст, x, 1)): Next
End Function


Public Function Симв(символ_для_перевода_в_адресную_строку As String)
    Select Case символ_для_перевода_в_адресную_строку
        Case "!": Симв = "%21":        Case """": Симв = "%22":      Case "#": Симв = "%23":       Case "$": Симв = "%24":       Case "%": Симв = "%25"
        Case "&": Симв = "%26":        Case "'": Симв = "%27":       Case "(": Симв = "%28":       Case ")": Симв = "%29":       Case "*": Симв = "*"
        Case "+": Симв = "%2B":        Case ",": Симв = "%2C":       Case "-": Симв = "-":         Case ".": Симв = ".":         Case "/": Симв = "%2F"
        Case "0": Симв = "0":          Case "1": Симв = "1":         Case "2": Симв = "2":         Case "3": Симв = "3":         Case "4": Симв = "4"
        Case "5": Симв = "5":          Case "6": Симв = "6":         Case "7": Симв = "7":         Case "8": Симв = "8":         Case "9": Симв = "9"
        Case ":": Симв = "%3A":        Case ";": Симв = "%3B":       Case "<": Симв = "%3C":       Case "=": Симв = "%3D":       Case ">": Симв = "%3E"
        Case "?": Симв = "%3F":        Case "@": Симв = "%40":       Case "A": Симв = "A":         Case "B": Симв = "B":         Case "C": Симв = "C"
        Case "D": Симв = "D":          Case "E": Симв = "E":         Case "F": Симв = "F":         Case "G": Симв = "G":         Case "H": Симв = "H"
        Case "I": Симв = "I":          Case "J": Симв = "J":         Case "K": Симв = "K":         Case "L": Симв = "L":         Case "M": Симв = "M"
        Case "N": Симв = "N":          Case "O": Симв = "O":         Case "P": Симв = "P":         Case "Q": Симв = "Q":         Case "R": Симв = "R"
        Case "S": Симв = "S":          Case "T": Симв = "T":         Case "U": Симв = "U":         Case "V": Симв = "V":         Case "W": Симв = "W"
        Case "X": Симв = "X":          Case "Y": Симв = "Y":         Case "Z": Симв = "Z":         Case "[": Симв = "%5B":       Case "\": Симв = "%5C"
        Case "]": Симв = "%5D":        Case "^": Симв = "%5E":       Case "_": Симв = "_":         Case "`": Симв = "%60":       Case "a": Симв = "a"
        Case "b": Симв = "b":          Case "c": Симв = "c":         Case "d": Симв = "d":         Case "e": Симв = "e":         Case "f": Симв = "f"
        Case "g": Симв = "g":          Case "h": Симв = "h":         Case "i": Симв = "i":         Case "j": Симв = "j":         Case "k": Симв = "k"
        Case "l": Симв = "l":          Case "m": Симв = "m":         Case "n": Симв = "n":         Case "o": Симв = "o":         Case "p": Симв = "p"
        Case "q": Симв = "q":          Case "r": Симв = "r":         Case "s": Симв = "s":         Case "t": Симв = "t":         Case "u": Симв = "u"
        Case "v": Симв = "v":          Case "w": Симв = "w":         Case "x": Симв = "x":         Case "y": Симв = "y":         Case "z": Симв = "z"
        Case "{": Симв = "%7B":        Case "|": Симв = "%7C":       Case "}": Симв = "%7D":       Case "~": Симв = "~":         Case "": Симв = "%7F"
        Case "Ђ": Симв = "%D0%82":     Case "Ѓ": Симв = "%D0%83":    Case "‚": Симв = "%E2%80%9A": Case "ѓ": Симв = "%D1%93":    Case """": Симв = "%E2%80%9E"
        Case "…": Симв = "%E2%80%A6":  Case "†": Симв = "%E2%80%A0": Case "‡": Симв = "%E2%80%A1": Case "?": Симв = "%E2%82%AC": Case "‰": Симв = "%E2%80%B0"
        Case "Љ": Симв = "%D0%89":     Case "‹": Симв = "%E2%80%B9": Case "Њ": Симв = "%D0%8A":    Case "Ќ": Симв = "%D0%8C":    Case "Ћ": Симв = "%D0%8B"
        Case "Џ": Симв = "%D0%8F":     Case "ђ": Симв = "%D1%92":    Case "'": Симв = "%E2%80%98": Case "'": Симв = "%E2%80%99": Case """": Симв = "%E2%80%9C"
        Case """": Симв = "%E2%80%9D": Case "o": Симв = "%E2%80%A2": Case "-": Симв = "%E2%80%93": Case "-": Симв = "%E2%80%94": Case "?": Симв = "%C2%98"
        Case "™": Симв = "%E2%84%A2":  Case "љ": Симв = "%D1%99":    Case "›": Симв = "%E2%80%BA": Case "њ": Симв = "%D1%9A":    Case "ќ": Симв = "%D1%9C"
        Case "ћ": Симв = "%D1%9B":     Case "џ": Симв = "%D1%9F":    Case " ": Симв = "%C2%A0":    Case "Ў": Симв = "%D0%8E":    Case "ў": Симв = "%D1%9E"
        Case "Ј": Симв = "%D0%88":     Case "¤": Симв = "%C2%A4":    Case "Ґ": Симв = "%D2%90":    Case "¦": Симв = "%C2%A6":    Case "§": Симв = "%C2%A7"
        Case "Ё": Симв = "%D0%81":     Case "©": Симв = "%C2%A9":    Case "Є": Симв = "%D0%84":    Case """": Симв = "%C2%AB":   Case "": Симв = "%C2%AC"
        Case "": Симв = "%C2%AD":      Case "®": Симв = "%C2%AE":    Case "Ї": Симв = "%D0%87":    Case "°": Симв = "%C2%B0":    Case "±": Симв = "%C2%B1"
        Case "І": Симв = "%D0%86":     Case "і": Симв = "%D1%96":    Case "ґ": Симв = "%D2%91":    Case "µ": Симв = "%C2%B5":    Case "": Симв = "%C2%B6"
        Case "·": Симв = "%C2%B7":     Case "ё": Симв = "%D1%91":    Case "№": Симв = "%E2%84%96": Case "є": Симв = "%D1%94":    Case """": Симв = "%C2%BB"
        Case "ј": Симв = "%D1%98":     Case "Ѕ": Симв = "%D0%85":    Case "ѕ": Симв = "%D1%95":    Case "ї": Симв = "%D1%97":    Case "А": Симв = "%D0%90"
        Case "Б": Симв = "%D0%91":     Case "В": Симв = "%D0%92":    Case "Г": Симв = "%D0%93":    Case "Д": Симв = "%D0%94":    Case "Е": Симв = "%D0%95"
        Case "Ж": Симв = "%D0%96":     Case "З": Симв = "%D0%97":    Case "И": Симв = "%D0%98":    Case "Й": Симв = "%D0%99":    Case "К": Симв = "%D0%9A"
        Case "Л": Симв = "%D0%9B":     Case "М": Симв = "%D0%9C":    Case "Н": Симв = "%D0%9D":    Case "О": Симв = "%D0%9E":    Case "П": Симв = "%D0%9F"
        Case "Р": Симв = "%D0%A0":     Case "С": Симв = "%D0%A1":    Case "Т": Симв = "%D0%A2":    Case "У": Симв = "%D0%A3":    Case "Ф": Симв = "%D0%A4"
        Case "Х": Симв = "%D0%A5":     Case "Ц": Симв = "%D0%A6":    Case "Ч": Симв = "%D0%A7":    Case "Ш": Симв = "%D0%A8":    Case "Щ": Симв = "%D0%A9"
        Case "Ъ": Симв = "%D0%AA":     Case "Ы": Симв = "%D0%AB":    Case "Ь": Симв = "%D0%AC":    Case "Э": Симв = "%D0%AD":    Case "Ю": Симв = "%D0%AE"
        Case "Я": Симв = "%D0%AF":     Case "а": Симв = "%D0%B0":    Case "б": Симв = "%D0%B1":    Case "в": Симв = "%D0%B2":    Case "г": Симв = "%D0%B3"
        Case "д": Симв = "%D0%B4":     Case "е": Симв = "%D0%B5":    Case "ж": Симв = "%D0%B6":    Case "з": Симв = "%D0%B7":    Case "и": Симв = "%D0%B8"
        Case "й": Симв = "%D0%B9":     Case "к": Симв = "%D0%BA":    Case "л": Симв = "%D0%BB":    Case "м": Симв = "%D0%BC":    Case "н": Симв = "%D0%BD"
        Case "о": Симв = "%D0%BE":     Case "п": Симв = "%D0%BF":    Case "р": Симв = "%D1%80":    Case "с": Симв = "%D1%81":    Case "т": Симв = "%D1%82"
        Case "у": Симв = "%D1%83":     Case "ф": Симв = "%D1%84":    Case "х": Симв = "%D1%85":    Case "ц": Симв = "%D1%86":    Case "ч": Симв = "%D1%87"
        Case "ш": Симв = "%D1%88":     Case "щ": Симв = "%D1%89":    Case "ъ": Симв = "%D1%8A":    Case "ы": Симв = "%D1%8B":    Case "ь": Симв = "%D1%8C"
        Case "э": Симв = "%D1%8D":     Case "ю": Симв = "%D1%8E":    Case "я": Симв = "%D1%8F"
    End Select
End Function



