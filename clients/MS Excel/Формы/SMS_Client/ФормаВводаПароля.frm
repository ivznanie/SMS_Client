VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���������������� 
   Caption         =   "�����������"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "����������������.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "����������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ��_Click()
    If ���������������������� And �������������� Then
        ����� = �����������������.Text: ������ = ���������������.Text
        
        �����������������.Text = "": ���������������.Text = ""
        
        ����������������.Hide
    End If
End Sub


Private Function ����������������������()
    On Error GoTo er
    
    Dim reg As Object, math As Object
    
    �����������������.Text = Trim(�����������������.Text)
    
    If �����������������.Text = "" Then ���������������������� = False: MsgBox "������ �� ����� ���� ������!", vbExclamation: Exit Function
    
    Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^79\d{9}$": Set math = reg.Execute(�����������������.Text) '����� �������� ������
    
    If Not math.Count = 1 Then ���������������������� = False: MsgBox "������ ������ ������ ���� �����: 79998887766", vbExclamation Else ���������������������� = True
    
    Exit Function
er:
    ���������������������� = True
End Function

Private Function ��������������()
    ���������������.Text = Trim(���������������.Text)

    If ���������������.Text = "" Then �������������� = False: MsgBox "������ �� ����� ���� ������!", vbExclamation Else �������������� = True
End Function


Private Sub ��������������_Click()
    On Error GoTo er
    
    If ���������������������� Then
        ����� = GetHTTPResponse(Replace(URL_Password, "{login}", Trim(�����������������.Text)))
        
        ������ = ���������������������(�����)
        
        If ������ <> "ok" Then MsgBox ������, vbCritical: Exit Sub Else MsgBox "SMS � ������� ���������� �� ��������� ����� +" & �����������������.Text, vbInformation
        
        ���������������.Text = ""
    End If
    
    Exit Sub
er:
    MsgBox "������ ��������� ������: " & p(2) & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Sub

