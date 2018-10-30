Set emailObj      = CreateObject("CDO.Message")

emailObj.From     = "lukasz.teska@bd.com"
emailObj.To       = "lukasz.teska@bd.com"

emailObj.Subject  = "Test CDO"
emailObj.TextBody = "Test CDO"

Set emailConfig = emailObj.Configuration

emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = true 
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")    = "YourUserName"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = "Password1"

emailConfig.Fields.Update

emailObj.Send

If err.number = 0 then Msgbox "Done"