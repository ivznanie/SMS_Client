# Сервис SMS_Client
(релиз 1.0.1. от 8 июня 2017 года)

## Общая информация
Сервис позволяет взаимодействовать с провайдерами услуг отправки СМС (SigmaSMS)
сообщений  через собственный универсальный API-интерфейс и через клиента в виде файла 
формата .xlsm

## Настройки 
+ Пути к точкам доступа API SigmaSMS передаюстя строкой формат json
+ Логины зарегистрированных пользователей по умолчанию в user.csv.
+ Хэши паролей по умолчанию в login.csv. Если файла нет, то создается при установке пароля первый раз.

## Запросы для обращения к API:
+ отправить смс: http://test.od37.ru/sms/sms/send/?login=79106679925&password=7cb23cd1&phone=79106679925&text=%D0%97%D0%B4%D0%BE%D1%80%D0%BE%D0%B2%D0%BE
+ получить статус отправленного сообщения: http://test.od37.ru/sms/sms/status/?login=79106679925&password=7cb23cd1&id=4953208191406954340001
+ новый пароль: http://test.od37.ru/sms/user/newpass/?login=79106679925
+ получить текущий баланс пользователя: http://test.od37.ru/sms/user/balance/?login=79106679925&password=7cb23cd1
+ версия (релиз): http://test.od37.ru/sms/version/
+ статус сервиса: http://test.od37.ru/sms/status/


## Ответы API:
###Ответы - успехи:
+ 0: SMS шлюз готов к принятию запросов
+ 1: SMS шлюз находится в режиме технического обслуживания
+ 100:Сообщение отправлено&id:5074094692618805620001&стоимость:1.1
+ 101:Cтатус сообщения&доставлено:2017-05-21 01-55-47 (или &недоставлено)
+ 103:Баланс пользователя в рублях&баланс:125.52
+ 104:Новый пароль установлен
+ 105:Новый пользователь зарегистрирован

###Ответы - ошибки:

####Группа "300" - ошибки, авторизации пользователя
+ 301:Пользователь не определен
+ 302:Пользователь не зарегистрирован
+ 303:Пользователь не авторизован
+ 304:Новый пароль установить не удалось
+ 305:Баланс получить не удалось

####Группа "310" - ошибки, параметров запроса
+ 311:Не указаны телефон получателя и/или текст сообщения
+ 312:Неверный формат телефона получателя
+ 313:Не указан текст сообщения
+ 314:Превышена допустимая длина сообщения
+ 315:Неверный id сообщения

####Группа "320" - ошибки, прав пользователя
+ 320:Не достаточно средств

####Группа "400" - ошибки сервиса
<<<<<<< HEAD
+ 400:Ошибка сервиса
=======
+ 400:Ошибка сервиса
>>>>>>> 74bb844baa9ae3f56ccae6651770122a0628f9d0
