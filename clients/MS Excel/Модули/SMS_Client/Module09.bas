Attribute VB_Name = "Module09"
'============================================================================================================================================================
'====================================================================���������������=========================================================================
'============================================================================================================================================================

Sub pr(��������, Optional �������� As Boolean = True)
    On Error Resume Next
    
    If �������� Then
        Select Case ��������
            Case "�����������":      Sheets("�����������").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True: Sheets("�����������").EnableSelection = xlNoRestrictions
            Case "���������":        Sheets("���������").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True: Sheets("���������").EnableSelection = xlUnlockedCells
            Case "������ ��������":  Sheets("������ ��������").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
        End Select
    Else
        Select Case ��������
            Case "�����������":      Sheets("�����������").Unprotect: Sheets("�����������").Range("a6").Select
            Case "���������":        Sheets("���������").Unprotect
            Case "������ ��������":  Sheets("������ ��������").Unprotect
        End Select
    End If
End Sub


Function p(Optional n = 1)
    On Error Resume Next
    
    p = "": For i = 1 To n: p = p & Chr(13) & Chr(10): Next
End Function

Function p2(Optional n = 1)
    On Error Resume Next
    
    p2 = "": For i = 1 To n: p2 = p2 & Chr(10): Next
End Function


Sub �����������������()
    On Error Resume Next
    
    If Not ApplicationUndo Then ApplicationUndo = True: Application.Undo
    
    ApplicationUndo = False
End Sub


Function gen(max)
    Randomize: gen = CInt(Int((32767 * Rnd()) + 1))
End Function


Function �����������������������(ByRef ����, Optional ������������ = 1)
    On Error Resume Next
    
    ����������������������� = ����.Cells(����.Rows.Count, ������������).End(xlUp).Row
End Function
