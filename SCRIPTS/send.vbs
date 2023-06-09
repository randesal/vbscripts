Set memail = CreateObject("CDO.Message")
memail.Subject = "Attachment"
memail.From = "rsalillas@net.tonghsing.ph"
memail.To = "rsalillas@net.tonghsing.ph"
memail.TextBody = ""
memail.AddAttachment "E:\DESKTOP\Script\for_send\K_File Generator.zip"
memail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
memail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.tonghsing.ph"
memail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
memail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
memail.Configuration.Fields.Update
memail.send
Set memail = Nothing