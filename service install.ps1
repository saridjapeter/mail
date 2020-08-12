#Set-ExecutionPolicy RemoteSigned
$NSSMPath = ($PSScriptRoot+ "\nssm.exe")
$NewServiceName = “ClaimDeskMailSend”
$PoShPath= (Get-Command powershell).Source
$PoShScriptPath = ($PSScriptRoot+ “\mail.ps1”)
$args = '-ExecutionPolicy Bypass -NoProfile -File "{0}"' -f $PoShScriptPath
& $NSSMPath install $NewServiceName $PoShPath $args
& $NSSMPath status $NewServiceName

Start-Service $NewServiceName

Get-Service $NewServiceName

#(Get-WmiObject win32_service -Filter "name='ClaimDeskMailSend'").delete()