On Error Resume Next

'###PEGAR INFO DO SHAREPOINT'
sharepoint = InputBox("Link do sharepoint")
usuario = InputBox("email")
senha = InputBox("senha")

Const HKEY_CURRENT_USER = &H80000001

'###PRIMEIRO DOMINIO DA MICROSOFT###'

strComputer = "."
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\Domains\*.sharepoint.com"

objReg.CreateKey HKEY_CURRENT_USER, strKeyPath

strValueName = "https"
dwValue = 2

objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

'###SEGUNDO DOMINIO DA MICROSOFT###'

strComputer = "."
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\Domains\*.microsoft.com"

objReg.CreateKey HKEY_CURRENT_USER, strKeyPath

strValueName = "https"
dwValue = 2

objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

'###TERCEIRO DOMINIO DA MICROSOFT###'

strComputer = "."
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\Domains\*.microsoftonline.com"

objReg.CreateKey HKEY_CURRENT_USER, strKeyPath

strValueName = "https"
dwValue = 2

objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

'###SALVANDO CREDENCIAL DA WEB###'

Set oShell = WScript.CreateObject("WSCript.shell")
oShell.run "cmd /c cmdkey /generic:https://login.microsoftonline.com /user:" & usuario & " /pass:" & senha
oShell.run "cmd /c cmdkey /generic:sharepoint /user:" & usuario & " /pass:" & senha

WScript.Echo "Entre com a conta da microsoft usada no office 365 e salve sua senha"

Dim Shell
Set Shell = CreateObject("WScript.Shell")
Shell.Run "iexplore.exe " & sharepoint

WScript.Echo "Feito?"
WScript.StdIn.ReadLine


'###MAPEANDO UNIDADE DE REDE###'

Set objRede = WScript.CreateObject("Wscript.Network")
Set objArq = CreateObject("Scripting.FileSystemObject")

driveX = "X:"
pathX = sharepoint
objRede.MapNetworkDrive driveX, pathX

'##FINALIZADO###'

WScript.Echo "Finalizado"

