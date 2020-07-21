#Install-Module -Name SqlServer
#Install-Module -Name SqlServer -Scope CurrentUser

#$ins = "192.168.0.11"
#$user="sa"
#$pass="pCOc6Ly79p"
#Берем настройки с dbo.Config
$config=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "set nocount on;SELECT Data FROM Config"
#Смотрим письмо в Email
$query=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "set nocount on;SELECT TOP (1) Receiver, Message_Text, Subject, ID_Attachment, ID_Source FROM Email"
#Добавляем в Dispatch
Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "INSERT INTO Dispatch (Receiver, Message, ID_Source, ID_Attach, Message_Type, Message_Time) SELECT TOP (1) Receiver, Message_Text, ID_Source, ID_Attachment, Message_Type='0', Message_Time=GETDATE()  FROM Email"  
#Удаляем запись в Email
#Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "DELETE FROM Email ORDER BY ID_Email ASC LIMIT 1"
#Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "DELETE TOP(1) FROM Email"


#Адрес сервера SMTP для отправки
$serverSmtp = $config[6].Data 

#Порт сервера
$port = $config[26].Data 

#От кого
$From = $config[5].Data 

#Кому
$To = $query[0]
#$To = "p.saridja@be2b.ru" 

#Тема письма
$subject = $query[2]

#Логин и пароль от ящики с которого отправляете login@yandex.ru
$user = $config[7].Data 
$pass = $config[8].Data 

#Путь до файла 
#$file = "D:\zabbix\tets.txt"

#Создаем два экземпляра класса
#$att = New-object Net.Mail.Attachment($file)
$mes = New-Object System.Net.Mail.MailMessage

#Формируем данные для отправки
$mes.From = $from
$mes.To.Add($to) 
$mes.Subject = $subject 
$mes.IsBodyHTML = $true 
$mes.Body = $query[1]

#Добавляем файл
#$mes.Attachments.Add($att) 

#Создаем экземпляр класса подключения к SMTP серверу 
$smtp = New-Object Net.Mail.SmtpClient($serverSmtp, $port)

#Сервер использует SSL 
$smtp.EnableSSL = $true 

#Создаем экземпляр класса для авторизации на сервере яндекса
$smtp.Credentials = New-Object System.Net.NetworkCredential($user, $pass);

#Отправляем письмо, освобождаем память
$smtp.Send($mes) 
#$att.Dispose()