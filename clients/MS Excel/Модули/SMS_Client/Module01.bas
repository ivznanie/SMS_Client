Attribute VB_Name = "Module01"
'======================================================================================================================================================
'============================������ ����������� ����� ��������, ����������, ������������� �����, ��������������� �����=================================
'======================================================================================================================================================


'===================================================================���������==========================================================================

Public Const ����������������������������������������� = "ERROR:" '����� ������, ������������

Public Const ���������������� = "��� ������ � ������!", ������������������������� = "����������� �������� ���������� �������..."

Public Const ����������������������������������� = "���� ���� ���������� ��������, ���������� ������������ ����� ������."
Public Const ������������������������������������ = "������������� ������������ ����� ������, ��� ��� ������� ����� �������� � �������������."
Public Const �������������������������������������� = "������������ ����� ������ ����� ������ ����� ��������� ����."

'===================================================================����������=========================================================================

Public URL_Version, URL_Status, URL_Password, URL_Balance, URL_SendSMS, URL_StatusSMS, URL_Root  'URL API ���-�������

Public ���������(), ����(), ����������(), ������, �����, �����, �������������������������������, ������������, ��������������������

Public ApplicationUndo, ���������������, �������������������������������VBA, ��������������������, ���������������, URL��������������������


'==================================================================�������������=======================================================================

Sub �������������() '��������� ������������� ����� ��� �������
    �������������_
End Sub


Sub �������������_()
    On Error GoTo er
        
    If Not �������������� Then Exit Sub Else �������������������������
    
    Call ��������������������������������: If Not ������������������(True) Then Exit Sub  '�������� ���������� ���-������� � ������
        
    Exit Sub
er:
    MsgBox "������ ��������� �������������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Sub

 
    Sub �������������������������()
        On Error Resume Next
        
        ���� = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\"
     
        Set WshShell = CreateObject("WScript.Shell"): If Err.Number <> 0 Then Exit Sub
        
        With WshShell
            �������� = .RegRead(���� & "VBAWarnings") '�������� ������ ������������ ��������
            
            If ��������������� And �������� <> "1" Then
'                If MsgBox("��������! ��� ������ �������� ���������� ��������� �� ���������� � Excel. ������� ���������� ��� ������ ������������ �����������. " & p(2) & _
'                          "������� ��, ���� ������ �� ���������� �� ���������� ��� ��� � ������ ������.", vbQuestion + vbYesNo) = vbYes Then

                    .RegWrite ���� & "VBAWarnings", 1, "REG_DWORD": If Err.Number = 0 Then �������������������������� = True Else ������ = True
'                End If
            End If
            
            �������� = .RegRead(���� & "AccessVBOM") '�������� ������������ ������� � ��������� ������ ������� VBA
         
            If �������������������������������VBA And �������� <> "1" Then
'                If MsgBox("��������! ��� ������ �������� ���������� ��������� ������ � ��������� ������ ������� VBA (Excel). ������� ���������� ��� ������ ������������ �����������. " & p(2) & _
'                          "������� ��, ���� ������ �� ���������� ������ ��� ��� � ������ ������.", vbQuestion + vbYesNo) = vbYes Then
            
                    .RegWrite ���� & "AccessVBOM", 1, "REG_DWORD": If Err.Number = 0 Then �������������������������� = True Else ������ = True
'                End If
            End If
            
            'If �������������������������� And Not ������ Then MsgBox "��� ���������� �������� ����� ����� �������, �������� �� ����� ��� ������.", vbInformation: ThisWorkbook.Close 0
        End With
        
        If ������ Then _
            MsgBox "��������! Excel ������� �������� ������� � �������� � ��������� ������ �������� ���������� �����, �� � ����, ������, ��� �� ����������. ����������� ������� ��� �������." & p(3) & _
                   "��� �����: " & p(2) & "1. ��������� � ���� Excel:" & p(1) & "    �) �������� ����� ���������� �������������" & p(1) & "    �) �������� ��������� ������ ���������� �������������" & p(1) & _
                   "    �) �������� ��������� ��������" & p(2) & "2. �������� �����: '�������� ��� �������' [��������� �����]" & p(2) & "3. ���������� �������: �������� ������ " & _
                   "� ��������� ������ �������� VBA" & p(2) & "4. �������� Excel � ��������� ���� ������", vbExclamation
    End Sub


    Function ������������������(Optional ����������� = False)
        On Error GoTo er
        
        If ����������� Then ���������������� �������������������������, "", 1
        
        If Not ����������������(���������������, �����������) Then Exit Function
                
        ������������������ = True: If ����������� Then ���������������� ����������������, "" '����� ��������� ����������� ���� ��������
        
        Exit Function
er:
        MsgBox "������ �������� ��������� �������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
    End Function
    
        
        Function ����������������(���������������������, �����������)
            On Error GoTo er
            
            If ��������������������� = "ok" Then: ���������������� = True: Exit Function
            
            Dim reg As Object, math As Object: Set reg = CreateObject("vbscript.regexp")

            ��������������� = 2 '������� �� ���������
            
            If InStr(���������������������, ������������������������������������) > 0 Or InStr(���������������������, ��������������������������������������) > 0 Then _
                ��������������� = 1 Else If InStr(���������������������, �����������������������������������) > 0 Then ��������������� = 2 '�������������
                
            reg.Pattern = "((http://)|(https://)|(ftp://)).+\.xlsm": Set math = reg.Execute(���������������������):  If math.Count = 1 Then URL = math(0).value

            If ����������� Then ���������������� ���������������������, URL, ��������������� Else If ��������������� = 2 Then MsgBox ���������������������, IIf(��������������� = 2, vbCritical, vbExclamation)
        
            If ��������������� <> 2 Then ���������������� = True
        
            Exit Function
er:
            MsgBox "������ �������� ������� ����������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
        End Function
 
 
        Function ���������������() As String
            On Error GoTo er
                
            Dim reg As Object, math As Object, i: Set reg = CreateObject("vbscript.regexp")
            
            '������ ������� WEB-�������
                '������ ������:   "0:SMS ���� ��������� �������"
                '������ ������:   "1:SMS ���� � ������ ������������ ������������"
                '������ ������:   "2.1:https://github.com/ivznanie/SMS_Client/raw/dev/clients/MS Excel/����������/SMS_Client/SMS_Client (������ �������� ������������ ���).xlsm:1.0.1.1&2.2:https://github.com/ivznanie/SMS_Client/raw/dev/clients/MS Excel/����������/SMS_Client/SMS_Client (������ �������� ��������� � ���).xlsm:1.0.1.0:16.03.2018"
            
            ������������� = Trim(GetHTTPResponse(URL_Status)): ������ = ���������������������(�������������): If ������ <> "ok" Then ��������������� = ������: Exit Function '������ �������
            
            If Not (Left(�������������, 1) = "0" Or Left(�������������, 1) = "1" Or Left(�������������, 1) = "2") Then ��������������� = "WEB-������ ����������� ������������ ������ ������  �������": Exit Function
        
            '�������� ������� ���.������������ WEB-�������
            
            If Left(�������������, 2) = "1:" Then ������������������������������ = Right(�������������, Len(�������������) - 2) & p(2) & "���������� ����� �����...": Exit Function '� ������ ���������������
            
            
            '�������� ���������� � ������������� ������ ������� � WEB-������� (1.2.3. = 1'.2'.3'.) � ��������� ������ �� �������������� ������� (��) � ������ ��������������
            
            '������ ������ WEB-�������
                '������ ������:   1.0.1.0&https://github.com/ivznanie/SMS_Client/raw/dev/clients/MS Excel/����������/SMS_Client/SMS_Client (������ �������� ������������ ���).xlsm&https://github.com/ivznanie/SMS_Client/raw/dev/clients/MS Excel/����������/SMS_Client/SMS_Client (������ �������� ��������� � ���).xlsm
            
            ����� = Trim(GetHTTPResponse(URL_Version)): ������ = ���������������������(�����): If ������ <> "ok" Then ��������������� = ������: Exit Function '������ �������
            
            ����������� = Split(�����, "&"): If UBound(�����������) - LBound(�����������) + 1 <> 3 Then ��������������� = "WEB-������ ����������� ������������ ������ ������  ������": Exit Function
            
            ������������� = �����������(0): URL�������������������� = Trim(�����������(����������))
           
            reg.Pattern = "^\d+\.\d+\.\d+\.\d+$": Set math = reg.Execute(�������������): If math.Count = 0 Then ��������������� = "WEB-������ ����������� ������������ ������ �������� ����� ������: [" & ������������� & "] ��������� � �������: �.�.�.�": Exit Function
            
            mst = Split(�������������, "."): �������������123 = mst(0) & "." & mst(1) & "." & mst(2)
            mst = Split(�������������, "."): �������������123 = mst(0) & "." & mst(1) & "." & mst(2)
                        
            If �������������123 <> �������������123 Then ��������������� = "������ ����� Excel �� ������������� ������ WEB-�������: ��������� ������ " & �������������123 & ".X " & IIf(URL�������������������� <> "", p(2) & "���� ��� ����������: " & URL��������������������, ""): Exit Function
                        
                        
            '�������� WEB-������� � ���������� ������ �������� ������� (��� �������� ������������� - ��������)
            
            If Left(�������������, 2) = "0:" Then ��������������� = "ok": Exit Function
                
                
            '�������� ������� WEB-������� �� ��������������� (�������������) ��� ����������� (���������) ���������� ������� (��)
            
            reg.Pattern = "^2\.1:.+&2\.2:.+$": Set math = reg.Execute(�������������): If math.Count = 0 Then ��������������� = "WEB-������ ����������� ������������ ������ �������": Exit Function
            
            ������������������� = Split(�������������, "&")(���������� - 1):  ��������������������� = Trim(Right(�������������������, Len(�������������������) - InStr(�������������������, ":")))
            
            URL�������������������� = Trim(Left(���������������������, InStr(���������������������, "xlsm") + 3)): ����������������������URL = Replace(���������������������, URL�������������������� & ":", "")
            
            �������������������������� = Trim(Split(����������������������URL, ":")(0)): If InStr(����������������������URL, ":") > 0 Then ��������������������� = Replace(Trim(Split(����������������������URL, ":")(1)), "-", ".")
            
            reg.Pattern = "^\d+\.\d+\.\d+\.\d+$":  Set math = reg.Execute(��������������������������): If math.Count = 0 Then ��������������� = "WEB-������ ����������� ������������ ������ ������ ����� Excel ��� ����������: [" & �������������������������� & "] ��������� � �������: �.�.�.�": Exit Function
            
            
            '���� ������ ��� ���������� �� ����� �������, �����
            If Not ������������������������(�������������, ��������������������������) Then ��������������� = "ok": Exit Function
                        
            
            ��������������� = "����� ���������� ����� Excel, ������ " & �������������������������� & p(2) & "���� ��� ����������: " & URL��������������������
            
            If ��������������������� <> "" Then
                If Not IsDate(���������������������) Then ��������������������� = ��������������� & p(2) & "���� ���������� ����� Excel � ������� �����������: [" & Split(����������������������URL, ":")(1) & "] ��������� � �������: ��.��.����": Exit Function
        
                '���� ���� ���������� ����������, ����� � ���������� � �������������� ����������
                If CDate(���������������������) < Now Then ��������������� = ��������������� & p(2) & �����������������������������������: Exit Function
            End If
                
            '���� ����� ������ ������� ������� ��������� ������ ������� - ����� � ���������� �� ������������ ���������� ����� ��������� ����, ����� ������� ���������������� ��������� �� ����������
            m_�������������������������� = Split(��������������������������, "."): ��������������������������123 = m_��������������������������(0) & "." & m_��������������������������(1) & "." & m_��������������������������(2)
            If �������������123 = ��������������������������123 Then ��������� = ������������������������������������ Else ��������� = ��������������������������������������
                
            If ��������� = �������������������������������������� And ��������������������� = "" Then ��������������� = "�� ������� ���� ���������� ����� ������ ����� Excel, ����� ������� ������������� ������� ������ ����� �����������": Exit Function
                
            ��������������� = ��������������� & p(2) & IIf(��������������������� <> "", "���� ���������� ���������� ����� Excel " & ��������������������� & p(2), "") & ���������: Exit Function
           
            ��������������� = "ok" '��������...
            
            Exit Function
er:
            MsgBox "������ �������� ������ ��� ������� �������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
        End Function
        
        Function ������������������������(������, �������������������)
            On Error Resume Next
            
            m_������ = Split(������, "."): m_������������������� = Split(�������������������, "."): If UBound(m_������) - LBound(m_������) <> UBound(m_�������������������) - LBound(m_�������������������) Then Exit Function
            
            flag = False: For i = LBound(m_������) To UBound(m_������): flag = IIf(CInt(m_�������������������(i)) > CInt(m_������(i)), True, flag): Next: ������������������������ = flag
        End Function


        Sub ����������������(�����, URL, Optional ������ = 0) '������ = 0 (��), 1 (��������������), 2 (������)
            On Error GoTo er
            
            If ����� = "" Then Exit Sub
            
            Select Case ������
                Case 0: ���� = RGB(0, 255, 0): Case 1: ���� = RGB(180, 180, 0): Case 2: ���� = RGB(255, 0, 0)
            End Select
            
            With Sheets("�����������")
                Set ������ = .Range("a6"): If Not ������.value = ������������������������� And ����� = ���������������� Then Exit Sub
                
                pr "�����������", 0
                
                ������.ClearContents: ������.Font.Underline = xlUnderlineStyleNone: If URL <> "" Then .Hyperlinks.Add Anchor:=������, Address:=URL
                
                ������.value = �����: ������.Font.Color = ����

                pr "�����������"
            End With
        
            Exit Sub
er:
            'if Err.Number<>0 then MsgBox "������ ��������� �������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbExclamation
        End Sub


'=================================================================���������������======================================================================

Sub ���������������() '��������� ��������������� ����� ��� ��������
    ���������������_
End Sub


Sub ���������������_() '��������� ��������������� ����� ��� ��������
    On Error Resume Next
    
    �������������������� = True
    
    Application.DisplayAlerts = False
    
    ������������������������ = "��������! ���� �� ������ ��� ��������� ��� �������, ������ ���������� �������� �������, ��� ����� �������� ���������. " & _
                               "���� �� ������ ������ �������������� ������� ������������ (��� ������� ������������ Excel) - �������� � �.1, ����� �������� � �.2:" & p2(2) & _
                               "1. ������� �� ������ '���������' � ������ �������������� ������� ������������" & p2(1) & _
                               "1.1. � ����������� ���� �������� ����� '�������� ��� ����������'" & p2(1) & _
                               "1.2. ������������� �����" & p2(2) & _
                               "2. ���������: ���� -> ����� ���������� ������������� -> ��������� ������ ���������� ������������� -> ��������� ��������" & p2(1) & _
                               "2.1. �������� �����: '�������� ��� �������' [��������� �����]" & p2(1) & _
                               "2.2. ���������� �������: '�������� ������ � ��������� ������ �������� VBA'" & p2(1) & _
                               "2.3. ������������� �����"
    
    ���������������� ������������������������, "", 2
                               
    ��������������������������������
    
    ThisWorkbook.Save

    Application.DisplayAlerts = True
End Sub





'====================================================================����������=========================================================================

Sub ��������������������������������()
    On Error Resume Next
    
    ������� = Format(Date, "dd.mm.yyyy"): ��������������� = �������: Sheets("������ ��������").Range("c3").value = �������
    
    Sheets("�����������").Activate
End Sub































