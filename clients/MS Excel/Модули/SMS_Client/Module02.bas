Attribute VB_Name = "Module02"
'======================================================================================================================================================
'================================������ ����������� ����� ��������� ����������� �������� SMS-��������� � �������=======================================
'======================================================================================================================================================


'=================================================================�������� SMS=========================================================================

Function ���������SMS(���������, ��������, ���, ����������) As String  '����� ������� �������� SMS � ���������� ������ ���������� �������� � ������� � ������
    On Error GoTo er
    
    ��������������������� = URLEncode(���������): If InStr(���������������������, "ERROR:") > 0 Then ���������SMS = ���������������������: Exit Function
    
    URL_SendSMS_ = Replace(URL_SendSMS, "{login}", �����):       URL_SendSMS_ = Replace(URL_SendSMS_, "{password}", ������)
    URL_SendSMS_ = Replace(URL_SendSMS_, "{phone}", ��������):   URL_SendSMS_ = Replace(URL_SendSMS_, "{text}", ���������������������)

    �����������������SMS = Trim(GetHTTPResponse(URL_SendSMS_ & IIf(������������, "&debug=true", "")))
    
    ������ = ���������������������(�����������������SMS): If ������ <> "ok" Then ������������������ = ������: ���������SMS = ������ Else ������������������ = Replace(�����������������SMS, "&", Chr(10))
    
    ������ = ����������������(������������������)
    
    If ������ = "ok" Then If ���������(�����������������SMS, 100) Then ID = Trim(Split(Split(�����������������SMS, "&")(1), ":")(1)) Else ID = "�� ������ ��� ������ (" & Left(Trim(�����), 3) & ")"
    
    ������ = �����������������������(���������, ��������, ���, ����������, ������, ������������������, ID) 'Unicode
    
    If ������ <> "" Then MsgBox ������ & p(2) & "��� ������� �������� SMS ������ ������ ���������: " & ������

    If ���������SMS = "" Then ���������SMS = "ok"
    
    Exit Function
er:
    ���������SMS = "������ �������� SMS: " & Err.Source & ": " & Err.Number & ": " & Err.Description & ": " & "URL: " & URL_SendSMS_
End Function


    Function �����������������������(���������, ��������, ���, ����������, ������, ������������������, ID)
        On Error GoTo er
        
        ����������������������� = �����������������������(Sheets("������ ��������")): If ����������������������� < 5 Then Exit Function
        
        pr "������ ��������", 0
        
        With Sheets("������ ��������")
            With .Range("a6").Offset(����������������������� - 5, 0)
                With .Range(.Row & ":" & .Row): .Font.Color = RGB(0, 0, 0): .Interior.TintAndShade = 0: End With
                
                .value = Format(Date, "dd.mm.yyyy"): .Offset(0, 1).value = Format(Time, "hh:mm:ss"): .Offset(0, 2).value = ��������: .Offset(0, 3).value = ���������
                .Offset(0, 4).value = ���: .Offset(0, 5).value = ����������: .Offset(0, 6).value = ������: .Offset(0, 7).value = ������������������: .Offset(0, 8).value = ID
            End With
        End With
        
        pr "������ ��������"
        
        Exit Function
er:
        ����������������������� = "�������� �������� � ������ �������� �� �������! ������: " & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description
        
        pr "������ ��������"
    End Function


Function �����������������������()
    On Error GoTo er
    
    If Not �������������������������������� Then Exit Function Else If Not ������������������� Then Exit Function
    
    ����������������������� = True
    
    Exit Function
er:
    MsgBox "������ ����� ���������� ��������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function �������������������()
    On Error GoTo er
    
    If ������ = "" Or ����� = "" Then ����������������.Show
    If ������ = "" Or ����� = "" Then MsgBox "����� �������� � ������ �� ����� ���� �������!", vbExclamation: Exit Function
    
    ������������������� = True
    
    Exit Function
er:
    MsgBox "������ �������� �����������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ��������������������������������()
    On Error GoTo er
    
    Dim reg As Object, math As Object, i
    
    ����������������������� = �����������������������(Sheets("������ ��������")): If ����������������������� < 5 Then Exit Function
        
    ������������ = False: Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^79\d{9}$" '����� �������� ������
    
    With Sheets("������ ��������")
        ReDim ���������(0): ReDim ������������������������(0): ReDim ����(0): ReDim ����������(0)
            
        For i = 6 To �����������������������
            With .Range("a1").Offset(i - 1, 0)
                �������� = Trim(.value): ����������������������� = .Font.Bold
                
                If �������� <> "" And ����������������������� Then
                    Set math = reg.Execute(��������)
                    
                    If math.Count = 1 Then
                        .Font.Color = RGB(0, 0, 0)
                        
                        If ���������(LBound(���������)) <> "" Then ReDim Preserve ���������(UBound(���������) + 1): ReDim Preserve ����(UBound(����) + 1): ReDim Preserve ����������(UBound(����������) + 1):
                    
                        ���������(UBound(���������)) = ��������: ����(UBound(����)) = Trim(.Offset(0, 1).value): ����������(UBound(����������)) = Trim(.Offset(0, 2).value)
                    Else
                        .Font.Color = RGB(255, 0, 0): ������������ = True
                    End If
                End If
            End With
        Next
    End With
    
    If ������������ Then: MsgBox "��������! � ������ �������� ������������ ������ �� ����������� �������. ������ ������ ������ ���� �����: 79998887766", vbExclamation: Sheets("������ ��������").Activate
    
    If ���������(LBound(���������)) = "" Then MsgBox "� ������ �������� �� ������ ����������� ��� ��������� ������!", vbExclamation: Sheets("������ ��������").Activate: Exit Function
    
    Set reg = Nothing: Set math = Nothing
    
    �������������������������������� = True
    
    Exit Function
er:
    MsgBox "������ ��������� ������� ������ �������" & IIf(i > 0, " � ������ " & i, "") & "!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ���������������������������()
    On Error GoTo er
    
    Dim reg As Object, math As Object, i
    
    Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^\d{2}\.\d{2}\.\d{4}$" '����� �������� ����
    
    Set math = reg.Execute(Trim(Sheets("������ ��������").Range("c3"))):  If math.Count <> 1 Then MsgBox "��������� ���� ������� �������� ������ ���� � �������: ��.��.����, ��������: " & Format(Date, "dd.mm.yyyy"), vbExclamation: Exit Function
    
    If Not IsDate(math(0)) Then MsgBox "��������� ���� ������� �� �����", vbCritical: Exit Function

    ��������������������������� = True
    
    Exit Function
er:
    MsgBox "������ �������� ���� ������� ��������" & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Sub �����������������������()
	If ������������������������������������������� (1) <> "::" Then Exit Sub

    �����������������������_
End Sub


Sub �����������������������_()
    On Error GoTo er
    
    If Not ��������������������������� Then Exit Sub Else If Not ������������������� Then Exit Sub
    
    If Not �������������� Then Exit Sub Else If Not ������������������ Then Exit Sub
    
    ����������������������� = �����������������������(Sheets("������ ��������")): If ����������������������� < 5 Then Exit Sub
        
    ������������ = RGB(255, 0, 0): ����������� = RGB(0, 0, 0)
    
    �������������� = RGB(255, 200, 200): ������������� = RGB(255, 255, 200): �������������� = RGB(200, 255, 200)
        
    pr "������ ��������", 0
    
    With Sheets("������ ��������")
        ���� = .Range("c3").value: ������� = 0
    
        URL_StatusSMS_ = Replace(URL_StatusSMS, "{login}", �����):  URL_StatusSMS_ = Replace(URL_StatusSMS_, "{password}", ������)
    
        For Each ������ In .Range("a6:a" & �����������������������)
            ������������� = False: ����������� = 0: ����������������������� = �������
            
            If IsDate(������.value) Then
                ID = .Range("i" & ������.Row).value
                    
                If ID <> "" And CDate(������.value) = CDate(����) And Not (������.Interior.Color = �������������� Or ������.Interior.Color = ��������������) Then
                    If IsNumeric(ID) Then
                        ����� = Trim(GetHTTPResponse(Replace(URL_StatusSMS_, "{id}", ID))): ������ = ���������������������(�����)
                        
                        If ������ = "ok" And ���������(�����, 101) Then
                            ������ = ��������������(�����): If ������ <> "" Then ������� = ������� + 1
                            
                            Select Case ������
                                Case "����������": ����������� = ��������������: Case "��������������": ����������� = �������������: Case "������������": ����������� = ��������������
                            End Select
                         Else: ������������� = True: End If
                    Else: ������������� = True: End If
                End If
            Else: ������������� = True: End If
            
            With .Range(������.Row & ":" & ������.Row)
                If Not ������������� Then
                    .Font.Color = �����������: If ����������� <> 0 And ����������������������� < ������� Then .Interior.Color = ����������� Else .Interior.TintAndShade = 0
                Else
                    .Font.Color = ������������: .Interior.TintAndShade = 0
                End If
            End With
        Next
        
        ��������� = "��� ��������� ���� SMS-��������� ��������� ������ ��������/���������� ��� " & ������� & " �����."
        
        If ������� = 0 Then MsgBox ��������� & p(2) & "��������� �������:" & p(1) & "    1) ������ SMS-��������� ��� �� ��������;" & p(1) & _
                                   "    2) ��������� ��� �������� �� ��������� ���� �� ����������;" & p(1) & "    3) � ������� ������� ������� " & _
                                   "������;" & p(1) & "    4) ������ �� ������ �������������� SMS-���������.", vbExclamation _
                       Else _
                            MsgBox ���������, vbInformation
    End With
    
    pr "������ ��������"
    
    Exit Sub
er:
    MsgBox "������ �������� ������� ��������" & ": " & Err.Number & ": " & Err.Description, vbCritical
    
    pr "������ ��������"
End Sub

    
    Function ��������������(�����)
        If InStr(�����, ":") > 0 Then
            ������ = Trim(Split(Split(�����, "&")(1), ":")(0)): ���� = Trim(Split(Split(�����, "&")(1), ":")(1))
        Else
            ������ = Trim(Split(�����, "&")(1))
        End If
        
        If LCase(������) = "����������" Then
            If ���� <> "" Then �������������� = "����������" Else �������������� = "��������������"
            
        ElseIf LCase(������) = "������������" Then
            �������������� = "������������"
        End If
    End Function


Function ����������������������������(������)
    On Error Resume Next
    
    �������������������� = "": If ������ = "" Then Exit Function

    �������������������������� = Split(�����������������������������������������, " ")

    For i = LBound(��������������������������) To UBound(��������������������������)
        If InStr(������, ��������������������������(i)) > 0 Then �������������������� = �������������������� & " " & ��������������������������(i)
    Next
    
    For x = 1 To Len(������)
        If ����(Mid(������, x, 1)) = "" Then �������������������� = �������������������� & " " & Mid(������, x, 1)
    Next
    
    If �������������������� <> "" Then MsgBox "� ������ ��������� ��������� ������������: " & ��������������������, vbExclamation: Exit Function
    
    ���������������������������� = True
End Function


Function ���������������������(�����)
    On Error Resume Next
    
    ��������������������� = "ok"
    
    If Trim(�����) = "" Then ��������������������� = "Web-������ ������ ������ �����"
    
    If InStr(�����, "ERROR: GetHTTPResponse ����������� � �������:") > 0 Then '������ ������� GetHTTPResponse
        ��������������������� = "���������� ������ ��������� ������� � ���-�������. " & _
                                "��������� ���������� � ����������." & p2(2) & �����
        
    ElseIf InStr(�����, "TIMEOUT") > 0 Then '������ �� ��������
        ��������������������� = "���-������ �������� ����������. ���������� �����." & p2(2) & �����
        
    ElseIf InStr(�����, "ERROR:") > 0 Or ���������(�����, 400, ">=") Then '������ ������� ������ �������
        ��������������������� = "������ ���-�������. ���������� � ��������������." & p2(2) & �����
        
    ElseIf ���������(�����, 300, ">=") Then '������ ������� ������ �������
        ��������������������� = �����
    
        If ���������(�����, 301, ">=") And ���������(�����, 304, "<=") Then ����� = "": ������ = ""
    End If
End Function


Function ���������(ByVal �����, ��� As Integer, Optional ��������� = "=") As Boolean
    On Error Resume Next
    
    '���� ������ ���������� ��������� � ���'��
    
    ��������� = False: ����� = Trim(�����): If ����� = "" Then Exit Function
    
    ���������������� = Split(�����, "&")
    
    For i = LBound(���������) To UBound(���������)
        ���������������� = Trim(����������������(i)): ����� = Split(����������������, ":")
        
        If UBound(�����) - LBound(�����) + 1 = 2 Then
            ������������ = Trim(�����(LBound(�����)))
            
            If IsNumeric(������������) Then
                Select Case ���������
                    Case "=":  ��������� = (CInt(������������) = ���):  Case ">=": ��������� = (CInt(������������) >= ���):  Case "<=": ��������� = (CInt(������������) <= ���)
                    Case ">":  ��������� = (CInt(������������) > ���):  Case "<":  ��������� = (CInt(������������) < ���)
                End Select
            End If
        End If
        
        If ��������� Then Exit For
    Next
End Function


Function ����������������(�����)
    On Error GoTo er
    
    Select Case Left(Trim(�����), 3)
        Case "100": ���������������� = Split(Split(�����, Chr(10))(0), ":")(1) & " ���������� " & �����������������������(Split(Split(�����, Chr(10))(2), ":")(1))
        
        Case Else: If IsNumeric(Left(Trim(�����), 3)) And InStr(�����, ":") > 0 Then ���������������� = Split(�����, ":")(1) Else ���������������� = �����
    End Select
    
    Exit Function
er:
    ���������������� = �����
End Function


Sub ������������(�����, Optional ����� = False)
    On Error Resume Next
    
    With Sheets("���")
        ����������������������� = �����������������������(Sheets("���")): If ����������������������� < 2 Then Exit Sub
        
        ����������������������� = ����������������������� - IIf(�����, 1, 0)
        
        With .Range("a1").Offset(�����������������������, 0)
            If ����� Then ����� = 3
            
            .Offset(0, 0 + �����).value = Format(Date, "dd.mm.yyyy"): .Offset(0, 1 + �����).value = Format(Time, "hh:mm:ss"): .Offset(0, 2 + �����).value = �����
        End With
    End With
End Sub



'====================================================================������============================================================================

Sub �������������() '������ �������
    If ������������������������������������������� (1) <> "::" Then Exit Sub

    �������������_
End Sub

    
Sub �������������_() '������ �������
    On Error GoTo er
    
    Sheets("������").Range("d3").value = "": If Not �������������� Then Exit Sub

    If Not ������������������� Then Exit Sub Else If Not ������������������ Then Exit Sub
    
    URL_Balance_ = Replace(URL_Balance, "{login}", �����): URL_Balance_ = Replace(URL_Balance_, "{password}", ������)
    
    ����������������������� = Trim(GetHTTPResponse(URL_Balance_))
    
    ������ = ���������������������(�����������������������): If ������ <> "ok" Then MsgBox "��������� ������:" & p(2) & ������: Exit Sub
    
    Sheets("������").Range("d3").value = �����������������������(Split(Split(�����������������������, "&")(1), ":")(1))
    
    Exit Sub
er:
    MsgBox "������ ������� �������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description & p(2) & "URL: " & URL_SendSMS_, vbCritical
End Sub

    
    Function �����������������������(�����)
        On Error GoTo er
    
        ����� = Replace(�����, ".", ",")
        
        ����� = "": If InStr(�����, ",") > 0 Then ����� = Split(�����, ",")(0) Else ����� = �����
        
        ������� = "": If InStr(�����, ",") > 0 Then ������� = Split(�����, ",")(1): If Len(�������) = 1 Then ������� = ������� & "0"
        
        ����������������������� = IIf(����� <> "", ����� & " ������ ", "") & IIf(������� <> "", ������� & " ������", "")
    
        Exit Function
er:
        MsgBox "������ ���������� ���������� �������� �������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
    End Function





