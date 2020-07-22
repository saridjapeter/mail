﻿#Install-Module -Name SqlServer
#Install-Module -Name SqlServer -Scope CurrentUser

#$ins = "192.168.0.11"
#$user="sa"
#$pass="pCOc6Ly79p"
#Start-Transcript -Path ("D:\mail\"+(Get-Date -Format d)+".log")
while($true) {
#Смотрим письмо в Email
$query=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim "set nocount on;SELECT TOP (1) Receiver, Message_Text, Subject, ID_Attachment, ID_Source FROM Email"
if ($query) {
#Берем настройки с dbo.Config
$config=Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim "set nocount on;SELECT Data FROM Config"

#Адрес сервера SMTP для отправки
$serverSmtp = $config[6].Data 

#Порт сервера
$port = $config[26].Data 

#От кого
$From = $config[5].Data 

#Кому
$To = $query[0]

#Тема письма
$subject = $query[2]

#Логин и пароль от ящики с которого отправляете login@yandex.ru
$user = $config[7].Data 
$pass = $config[8].Data 

#Создаем два экземпляра класса
$mes = New-Object System.Net.Mail.MailMessage

#Формируем данные для отправки
$mes.From = $from
$mes.To.Add($to) 
$mes.Subject = $subject 
$mes.IsBodyHTML = $true 
$mes.Body = $query[1]

#Создаем экземпляр класса подключения к SMTP серверу 
$smtp = New-Object Net.Mail.SmtpClient($serverSmtp, $port)

#Сервер использует SSL 
$smtp.EnableSSL = $true 

#Создаем экземпляр класса для авторизации на сервере яндекса
$smtp.Credentials = New-Object System.Net.NetworkCredential($user, $pass);

#Отправляем письмо
$smtp.Send($mes) 

if ($?) { # Если SMTP сообщение было отправлено успешно, то
        write-host "Сообщение успешно отправлено"
        Add-content ("D:\mail\"+(Get-Date -Format d)+".log") (((Get-Date -Format u).Substring(0,(Get-Date -Format u).Length -1))+" Сообщение успешно отправлено")
        #Добавляем в Dispatch
        Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim "INSERT INTO Dispatch (Receiver, Message, ID_Source, ID_Attach, Message_Type, Message_Time) SELECT TOP (1) Receiver, Message_Text, ID_Source, ID_Attachment, Message_Type='0', Message_Time=GETDATE()  FROM Email"  
        #Удаляем запись в Email
        Invoke-Sqlcmd -ServerInstance 192.168.0.11 -Username sa -Password pCOc6Ly79p -Database Claim "DELETE TOP(1) FROM Email"
    } else { #Если SMTP сообщение не удалось отправить, то 
        write-host "Сообщение не отправлено. Ошибка:" $Error[0].ToString()
        Add-content ("D:\mail\"+(Get-Date -Format d)+".log") (((Get-Date -Format u).Substring(0,(Get-Date -Format u).Length -1))+ $Error[0].ToString())
    }

#Отправляем сообщение QUIT на SMTP-сервер (правильно завершаем TCP-подключение и освобождаем все ресурсы, используемые текущим экземпляром класса SmtpClient)
$mes.Dispose()
}
sleep 70   
}
#Stop-Transcript