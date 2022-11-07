
<%
Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail")
Mailer.FromAddress= Request.Form ("email")
Mailer.RemoteHost = "209.11.45.5"
Mailer.AddRecipient "Andrew Britton", "Andrew@AndrewBrittonBooks.com"
Mailer.Subject    = Request.Form ("subject")
Mailer.BodyText  = Request.Form ("message")
if Mailer.SendMail then
  Response.Write "<font face='verdana, arial' size='2' color='#000000'><div align='center'><b>Thank you.  Your message has been sent.</b> <br><br><br><a href='http://www.AndrewBrittonBooks.com'>back to AndrewBrittonBooks.com</a></div></font>"
else
  Response.Write "Mail send failure. Error was " & Mailer.Response
end if

%>
