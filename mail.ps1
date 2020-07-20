#& "C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\sqlcmd" -S 192.168.0.11 -U sa -P pCOc6Ly79p -Q "set nocount on;SELECT TOP (1) ID_Email, Receiver, Message_Text, Subject FROM Claim.dbo.Email" -h -1 -f 65001 -o D:\mail\mail.txt
##$a
#Install-Module -Name SqlServer
#Install-Module -Name SqlServer -Scope CurrentUser

#$ins = "192.168.0.11"
#$user="sa"
#$pass="pCOc6Ly79p"
$config=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "set nocount on;SELECT Data FROM Config"
$config[6]
$query=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim -Query "set nocount on;SELECT TOP (1) Receiver, Message_Text, Subject FROM Email"


#Адрес сервера SMTP для отправки
$serverSmtp = "$config[6]" 

#Порт сервера
$port = 587

#От кого
$From = "claimdesk_0@be2b.ru" 

#Кому
$To = "p.saridja@be2b.ru" 

#Тема письма
$subject = $query[2]

#Логин и пароль от ящики с которого отправляете login@yandex.ru
$user = "claimdesk_0@be2b.ru"
$pass = ""

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