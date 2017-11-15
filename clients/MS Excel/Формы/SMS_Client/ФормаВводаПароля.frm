VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ФормаВводаПароля 
   Caption         =   "Авторизация"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "ФормаВводаПароля.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ФормаВводаПароля"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ОК_Click()
    If ПроверкаНомераТелефона And ПроверкаПароля Then
        Логин = ОкноВводаТелефона.Text: Пароль = ОкноВводаПароля.Text
        
        ОкноВводаТелефона.Text = "": ОкноВводаПароля.Text = ""
        
        ФормаВводаПароля.Hide
    End If
End Sub


Private Function ПроверкаНомераТелефона()
    On Error GoTo er
    
    Dim reg As Object, math As Object
    
    ОкноВводаТелефона.Text = Trim(ОкноВводаТелефона.Text)
    
    If ОкноВводаТелефона.Text = "" Then ПроверкаНомераТелефона = False: MsgBox "Пароль не может быть пустым!", vbExclamation: Exit Function
    
    Set reg = CreateObject("vbscript.regexp"): reg.Pattern = "^79\d{9}$": Set math = reg.Execute(ОкноВводаТелефона.Text) 'маска проверки номера
    
    If Not math.Count = 1 Then ПроверкаНомераТелефона = False: MsgBox "Формат номера должен быть таким: 79998887766", vbExclamation Else ПроверкаНомераТелефона = True
    
    Exit Function
er:
    ПроверкаНомераТелефона = True
End Function

Private Function ПроверкаПароля()
    ОкноВводаПароля.Text = Trim(ОкноВводаПароля.Text)

    If ОкноВводаПароля.Text = "" Then ПроверкаПароля = False: MsgBox "Пароль не может быть пустым!", vbExclamation Else ПроверкаПароля = True
End Function


Private Sub ПолучитьПароль_Click()
    On Error GoTo er
    
    If ПроверкаНомераТелефона Then
        Ответ = GetHTTPResponse(Replace(URL_Password, "{login}", Trim(ОкноВводаТелефона.Text)))
        
        Ошибка = ПроверкаОшибкиЗапроса(Ответ)
        
        If Ошибка <> "ok" Then MsgBox Ошибка, vbCritical: Exit Sub Else MsgBox "SMS с паролем отправлена на указанный номер +" & ОкноВводаТелефона.Text, vbInformation
        
        ОкноВводаПароля.Text = ""
    End If
    
    Exit Sub
er:
    MsgBox "Ошибка получения пароля: " & p(2) & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
End Sub

