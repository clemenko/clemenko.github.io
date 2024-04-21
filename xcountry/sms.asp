<%
to_email = "4102123200@messaging.sprintpcs.com"
' change to address of your own SMTP server
	strHost = "iis.take-one.net"
		
		Set Mail = Server.CreateObject("Persits.MailSender")
		' enter valid SMTP host
		Mail.Host = strHost
		Mail.From = request.form("email") ' From address
		Mail.AddAddress to_email
		
		' message subject
		Mail.Subject = "from web"
		' message body
		
    Mail.Body = request.form("message")
		
	'End If	
		
		strErr = ""
		bSuccess = False
		On Error Resume Next ' catch errors
		Mail.Send	' send message
		If Err <> 0 Then ' error occurred
			strErr = Err.Description
		else
			bSuccess = True
		End If
	
	response.redirect("http://66.223.118.199/xcountry")

%>