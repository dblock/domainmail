%WINDIR%\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe Microsoft.DomainMail.dll /codebase
cscript smtpreg.vbs /add 1 OnArrival "Microsoft.DomainMail" Microsoft.DomainMail.Sink  "mail from=*"
