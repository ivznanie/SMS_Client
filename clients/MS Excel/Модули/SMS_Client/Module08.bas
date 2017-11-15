Attribute VB_Name = "Module08"
'======================================================================================================================================================
'======================================������ ����������� ��������������� ��������� �������� SMS-���������=============================================
'======================================================================================================================================================


'==========================================================���������� HTTP-��������====================================================================

Function GetHTTPResponse(ByVal sURL As String, Optional sParam As String = "") As String '����������� �����
    On Error GoTo er
    
    Const TIMEOUT& = 10  '� ��������
    
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    If sParam <> "" Then sParam = sParam & "&randomval=" & gen(32767) Else sURL = sURL & IIf(InStr(sURL, "?") > 0, "&", "?") & "randomval=" & gen(32767)
 
    ������������ sURL & IIf(sParam <> "", "?" & sParam, "")
    
    If sParam = "" Then
        winHttpReq.Open "GET", sURL, True: DoEvents: winHttpReq.Send: DoEvents
    Else
        winHttpReq.Open "POST", sURL, True: DoEvents
        winHttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        winHttpReq.Send sParam: DoEvents
    End If
 
    If Not winHttpReq.WaitForResponse(TIMEOUT&) Then GetHTTPResponse = "ERROR: GetHTTPResponse: ����� �������� �����, TIMEOUT = " & TIMEOUT&: Exit Function
    If winHttpReq.Status <> "200" Then GetHTTPResponse = "ERROR: GetHTTPResponse: " & winHttpReq.Status & ": " & winHttpReq.StatusText: Exit Function
     
    GetHTTPResponse = winHttpReq.responsetext
    
    ������������ GetHTTPResponse, True
    
    Set winHttpReq = Nothing
    
    Exit Function
er:
    GetHTTPResponse = "ERROR: GetHTTPResponse ����������� � �������: " & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description
    
    ������������ "����� �������� �� �������: " & GetHTTPResponse, True
End Function


'============================================================����������� URLENCODE======================================================================

Public Function URLEncode(ByVal ����� As String) '�������������� ��������� � ������������ ������ ��� ����������� �������� ����� ��������
    On Error Resume Next
    
    For x = 1 To Len(�����): URLEncode = URLEncode & ����(Mid(�����, x, 1)): Next
End Function


Public Function ����(������_���_��������_�_��������_������ As String)
    Select Case ������_���_��������_�_��������_������
        Case "!": ���� = "%21":        Case """": ���� = "%22":      Case "#": ���� = "%23":       Case "$": ���� = "%24":       Case "%": ���� = "%25"
        Case "&": ���� = "%26":        Case "'": ���� = "%27":       Case "(": ���� = "%28":       Case ")": ���� = "%29":       Case "*": ���� = "*"
        Case "+": ���� = "%2B":        Case ",": ���� = "%2C":       Case "-": ���� = "-":         Case ".": ���� = ".":         Case "/": ���� = "%2F"
        Case "0": ���� = "0":          Case "1": ���� = "1":         Case "2": ���� = "2":         Case "3": ���� = "3":         Case "4": ���� = "4"
        Case "5": ���� = "5":          Case "6": ���� = "6":         Case "7": ���� = "7":         Case "8": ���� = "8":         Case "9": ���� = "9"
        Case ":": ���� = "%3A":        Case ";": ���� = "%3B":       Case "<": ���� = "%3C":       Case "=": ���� = "%3D":       Case ">": ���� = "%3E"
        Case "?": ���� = "%3F":        Case "@": ���� = "%40":       Case "A": ���� = "A":         Case "B": ���� = "B":         Case "C": ���� = "C"
        Case "D": ���� = "D":          Case "E": ���� = "E":         Case "F": ���� = "F":         Case "G": ���� = "G":         Case "H": ���� = "H"
        Case "I": ���� = "I":          Case "J": ���� = "J":         Case "K": ���� = "K":         Case "L": ���� = "L":         Case "M": ���� = "M"
        Case "N": ���� = "N":          Case "O": ���� = "O":         Case "P": ���� = "P":         Case "Q": ���� = "Q":         Case "R": ���� = "R"
        Case "S": ���� = "S":          Case "T": ���� = "T":         Case "U": ���� = "U":         Case "V": ���� = "V":         Case "W": ���� = "W"
        Case "X": ���� = "X":          Case "Y": ���� = "Y":         Case "Z": ���� = "Z":         Case "[": ���� = "%5B":       Case "\": ���� = "%5C"
        Case "]": ���� = "%5D":        Case "^": ���� = "%5E":       Case "_": ���� = "_":         Case "`": ���� = "%60":       Case "a": ���� = "a"
        Case "b": ���� = "b":          Case "c": ���� = "c":         Case "d": ���� = "d":         Case "e": ���� = "e":         Case "f": ���� = "f"
        Case "g": ���� = "g":          Case "h": ���� = "h":         Case "i": ���� = "i":         Case "j": ���� = "j":         Case "k": ���� = "k"
        Case "l": ���� = "l":          Case "m": ���� = "m":         Case "n": ���� = "n":         Case "o": ���� = "o":         Case "p": ���� = "p"
        Case "q": ���� = "q":          Case "r": ���� = "r":         Case "s": ���� = "s":         Case "t": ���� = "t":         Case "u": ���� = "u"
        Case "v": ���� = "v":          Case "w": ���� = "w":         Case "x": ���� = "x":         Case "y": ���� = "y":         Case "z": ���� = "z"
        Case "{": ���� = "%7B":        Case "|": ���� = "%7C":       Case "}": ���� = "%7D":       Case "~": ���� = "~":         Case "": ���� = "%7F"
        Case "�": ���� = "%D0%82":     Case "�": ���� = "%D0%83":    Case "�": ���� = "%E2%80%9A": Case "�": ���� = "%D1%93":    Case """": ���� = "%E2%80%9E"
        Case "�": ���� = "%E2%80%A6":  Case "�": ���� = "%E2%80%A0": Case "�": ���� = "%E2%80%A1": Case "?": ���� = "%E2%82%AC": Case "�": ���� = "%E2%80%B0"
        Case "�": ���� = "%D0%89":     Case "�": ���� = "%E2%80%B9": Case "�": ���� = "%D0%8A":    Case "�": ���� = "%D0%8C":    Case "�": ���� = "%D0%8B"
        Case "�": ���� = "%D0%8F":     Case "�": ���� = "%D1%92":    Case "'": ���� = "%E2%80%98": Case "'": ���� = "%E2%80%99": Case """": ���� = "%E2%80%9C"
        Case """": ���� = "%E2%80%9D": Case "o": ���� = "%E2%80%A2": Case "-": ���� = "%E2%80%93": Case "-": ���� = "%E2%80%94": Case "?": ���� = "%C2%98"
        Case "�": ���� = "%E2%84%A2":  Case "�": ���� = "%D1%99":    Case "�": ���� = "%E2%80%BA": Case "�": ���� = "%D1%9A":    Case "�": ���� = "%D1%9C"
        Case "�": ���� = "%D1%9B":     Case "�": ���� = "%D1%9F":    Case " ": ���� = "%C2%A0":    Case "�": ���� = "%D0%8E":    Case "�": ���� = "%D1%9E"
        Case "�": ���� = "%D0%88":     Case "�": ���� = "%C2%A4":    Case "�": ���� = "%D2%90":    Case "�": ���� = "%C2%A6":    Case "�": ���� = "%C2%A7"
        Case "�": ���� = "%D0%81":     Case "�": ���� = "%C2%A9":    Case "�": ���� = "%D0%84":    Case """": ���� = "%C2%AB":   Case "": ���� = "%C2%AC"
        Case "": ���� = "%C2%AD":      Case "�": ���� = "%C2%AE":    Case "�": ���� = "%D0%87":    Case "�": ���� = "%C2%B0":    Case "�": ���� = "%C2%B1"
        Case "�": ���� = "%D0%86":     Case "�": ���� = "%D1%96":    Case "�": ���� = "%D2%91":    Case "�": ���� = "%C2%B5":    Case "": ���� = "%C2%B6"
        Case "�": ���� = "%C2%B7":     Case "�": ���� = "%D1%91":    Case "�": ���� = "%E2%84%96": Case "�": ���� = "%D1%94":    Case """": ���� = "%C2%BB"
        Case "�": ���� = "%D1%98":     Case "�": ���� = "%D0%85":    Case "�": ���� = "%D1%95":    Case "�": ���� = "%D1%97":    Case "�": ���� = "%D0%90"
        Case "�": ���� = "%D0%91":     Case "�": ���� = "%D0%92":    Case "�": ���� = "%D0%93":    Case "�": ���� = "%D0%94":    Case "�": ���� = "%D0%95"
        Case "�": ���� = "%D0%96":     Case "�": ���� = "%D0%97":    Case "�": ���� = "%D0%98":    Case "�": ���� = "%D0%99":    Case "�": ���� = "%D0%9A"
        Case "�": ���� = "%D0%9B":     Case "�": ���� = "%D0%9C":    Case "�": ���� = "%D0%9D":    Case "�": ���� = "%D0%9E":    Case "�": ���� = "%D0%9F"
        Case "�": ���� = "%D0%A0":     Case "�": ���� = "%D0%A1":    Case "�": ���� = "%D0%A2":    Case "�": ���� = "%D0%A3":    Case "�": ���� = "%D0%A4"
        Case "�": ���� = "%D0%A5":     Case "�": ���� = "%D0%A6":    Case "�": ���� = "%D0%A7":    Case "�": ���� = "%D0%A8":    Case "�": ���� = "%D0%A9"
        Case "�": ���� = "%D0%AA":     Case "�": ���� = "%D0%AB":    Case "�": ���� = "%D0%AC":    Case "�": ���� = "%D0%AD":    Case "�": ���� = "%D0%AE"
        Case "�": ���� = "%D0%AF":     Case "�": ���� = "%D0%B0":    Case "�": ���� = "%D0%B1":    Case "�": ���� = "%D0%B2":    Case "�": ���� = "%D0%B3"
        Case "�": ���� = "%D0%B4":     Case "�": ���� = "%D0%B5":    Case "�": ���� = "%D0%B6":    Case "�": ���� = "%D0%B7":    Case "�": ���� = "%D0%B8"
        Case "�": ���� = "%D0%B9":     Case "�": ���� = "%D0%BA":    Case "�": ���� = "%D0%BB":    Case "�": ���� = "%D0%BC":    Case "�": ���� = "%D0%BD"
        Case "�": ���� = "%D0%BE":     Case "�": ���� = "%D0%BF":    Case "�": ���� = "%D1%80":    Case "�": ���� = "%D1%81":    Case "�": ���� = "%D1%82"
        Case "�": ���� = "%D1%83":     Case "�": ���� = "%D1%84":    Case "�": ���� = "%D1%85":    Case "�": ���� = "%D1%86":    Case "�": ���� = "%D1%87"
        Case "�": ���� = "%D1%88":     Case "�": ���� = "%D1%89":    Case "�": ���� = "%D1%8A":    Case "�": ���� = "%D1%8B":    Case "�": ���� = "%D1%8C"
        Case "�": ���� = "%D1%8D":     Case "�": ���� = "%D1%8E":    Case "�": ���� = "%D1%8F"
    End Select
End Function



