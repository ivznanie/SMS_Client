Attribute VB_Name = "Module09"
'============================================================================================================================================================
'====================================================================ВСПОМОГАТЕЛЬНОЕ=========================================================================
'============================================================================================================================================================

Sub pr(ИмяЛиста, Optional Защитить As Boolean = True)
    On Error Resume Next
    
    If Защитить Then
        Select Case ИмяЛиста
            Case "Приветствие":      Sheets("Приветствие").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True: Sheets("Приветствие").EnableSelection = xlNoRestrictions
            Case "Настройки":        Sheets("Настройки").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True: Sheets("Настройки").EnableSelection = xlUnlockedCells
            Case "Журнал рассылки":  Sheets("Журнал рассылки").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
        End Select
    Else
        Select Case ИмяЛиста
            Case "Приветствие":      Sheets("Приветствие").Unprotect: Sheets("Приветствие").Range("a6").Select
            Case "Настройки":        Sheets("Настройки").Unprotect
            Case "Журнал рассылки":  Sheets("Журнал рассылки").Unprotect
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


Sub ОтменитьИзменение()
    On Error Resume Next
    
    If Not ApplicationUndo Then ApplicationUndo = True: Application.Undo
    
    ApplicationUndo = False
End Sub


Function gen(max)
    Randomize: gen = CInt(Int((32767 * Rnd()) + 1))
End Function


Function НижняяЗаполненнаяЯчейка(ByRef Лист, Optional НомерСтолбца = 1)
    On Error Resume Next
    
    НижняяЗаполненнаяЯчейка = Лист.Cells(Лист.Rows.Count, НомерСтолбца).End(xlUp).Row
End Function
