Attribute VB_Name = "Module11"
'======================================================================================================================================================
'===============================������ ����������� ������� ��������� ����������� �������� SMS-��������� � �������======================================
'======================================================================================================================================================


'===================================================================���������==========================================================================

Public Const ������������� = "1.0.1.0" '���������� ����������� ������ (�.�.�.�):  �1.�1.�1 = �2.�2.�2,  �1 = | <> �2

Public Const ���������� = "1"

Public Const ������������� = "Module01,Module02,Module08,Module09,Module11"
Public Const ������������ =  "�����������,������ ��������,SMS ��������,������ ��������,������,���,���������"
Public Const ���������� =    "����������������"

'======================================================================================================================================================

'===================================================================����������=========================================================================
'======================================================================================================================================================


Sub SMS_��������()
    If ������������������������������������������� (1) <> "::" Then Exit Sub

    SMS_��������_
End Sub


Sub SMS_��������_()
    On Error Resume Next
    
    If Not �������������� Then Exit Sub Else If Not ������������������ Then Exit Sub
    
    ��������� = Trim(Sheets("SMS ��������").Range("f4").value)
    
    If Not (����������������������� And �������������������������(���������)) Then Exit Sub
    
    '�������� SMS
    
    For i = LBound(���������) To UBound(���������)
        ����� = "": ����� = ���������SMS(���������, ���������(i), ����(i), ����������(i)): If ����� <> "ok" Then ����������� = ����������� + 1
    Next
        
    �������� = UBound(���������) - LBound(���������) + 1
    
    MsgBox IIf(����������� < ��������, "����� ������� ������������ SMS: " & CStr(�������� - �����������) & p, "") & _
           IIf(����������� > 0, "����� �� ������������ SMS: " & ����������� & p, ""), _
           IIf(����������� = 0, vbInformation, IIf(����������� = ��������, vbCritical, vbExclamation))
    
    Sheets("������ ��������").Activate
End Sub


Function �������������������������(���������)
    On Error GoTo er
    
    '�������� ������� ��������
    
    If ��������� = "" Then MsgBox "����� ��������� �����������!", vbExclamation: Exit Function
    
    If Not ����������������������������(���������) Then Exit Function
    
    ������������������������� = True
    
    Exit Function
er:
    MsgBox "������ ������� ���������� ��������!" & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ��������������()
    On Error GoTo er
	
    With Sheets("���������") '������ ��������
        URL_Root = Trim(.Range("c49").value)
        
        URL_Status = ���������������(.Range("c3").value):   URL_Version = ���������������(.Range("c4").value):  URL_Password = ���������������(.Range("c5").value)
        URL_Balance = ���������������(.Range("c6").value):  URL_SendSMS = ���������������(.Range("c7").value):  URL_StatusSMS = ���������������(.Range("c8").value)
        
        ������������������������������� = .Range("a50").value: ������������ = .Range("a51").value
        
        ��������������� = .Range("a53").value: �������������������������������VBA = .Range("a54").value:
    End With
   
    �������������� = True
    
    Exit Function
er:
    MsgBox "������ ������ ��������!" & ": "& Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Function


Function ���������������(��������)
    ��������������� = Replace(Trim(��������), "{urlroot}", URL_Root)
End Function






