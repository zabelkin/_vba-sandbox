Dim objOutlookApp As Object, objMail As Object

Set objOutlookApp = CreateObject("Outlook.Application")
objOutlookApp.Session.Logon
Set objMail = objOutlookApp.CreateItem(0)   'создаем новое сообщение
'если не получилось создать приложение или экземпляр сообщения - выходим
If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub

'    sTo = "AddressTo@mail.ru"    'Кому(можно заменить значением из ячейки - sTo = Range("A1").Value)
'    sSubject = "Автоотправка"    'Тема письма(можно заменить значением из ячейки - sSubject = Range("A2").Value)
'    sBody = "Привет от Excel-VBA"    'Текст письма(можно заменить значением из ячейки - sBody = Range("A3").Value)
'    sAttachment = "C:\Temp\Книга1.xls"    'Вложение(полный путь к файлу. Можно заменить значением из ячейки - sAttachment = Range("A4").Value)

'создаем сообщение
With objMail
	.To = sTo 'адрес получателя
	.CC = "" 'адрес для копии
	.BCC = "" 'адрес для скрытой копии
	.Subject = sSubject 'тема сообщения
	.BodyFormat = olFormatTXT
	.Body = sBody 'текст сообщения
'        .htmlBody = sBody 'текст сообщения в формате HTML
	.Attachments.Add sFName ' общее вложение с текстом письма
	.Attachments.Add sAttachment 'чтобы отправить активную книгу вместо sAttachment указать ActiveWorkbook.FullName'
'        .Attachments.Add "e:\Docs\Access\image002.jpg" ' общее вложение
'        .Send 'Send/Display, если необходимо просмотреть сообщение, а не отправлять без просмотра
	If IsSend Then 'если надо сразу отправить письмо
		.Send
	Else
		.Display   'если надо сначала посмотреть результат
	End If
End With

Set objOutlookApp = Nothing: Set objMail = Nothing