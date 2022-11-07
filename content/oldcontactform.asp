<!-- ########################################## -->
<!-- ########################################## -->
<!-- ######	  CAPTCHA SUBMISSION FORM  ######## -->
<!-- ########################################## -->
<!-- ########################################## -->

<% if (submit) = "yes" then %>
<!-- SEND/RESUBMIT -->

<% 
cTemp = recaptcha_confirm(recaptcha_private_key, recaptcha_challenge_field, recaptcha_response_field)
If cTemp <> "" Then 
%>

  <p><b>An error occured with the captcha. Please try again.</b></p>

<form action="contact.asp?submit=yes" method="post">
  <table cellspacing="2" cellpadding="0" border="0" width="490">
    <tr>
      <td align="left" width="140">First Name:<br /><input name="contact_firstname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_firstname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="left">Last Name:<br /><input name="contact_lastname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_lastname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="left">Email Address:<br /><input name="contact_email" style=" width: 200px;" class="textSize2 black" type="text" value="<%=contact_email%>" maxlength="150"></td>
    </tr>
    <tr>
      <td align="left" valign="top">Human Check:<br />
	  <%=recaptcha_challenge_writer(recaptcha_public_key)%></td>
      
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td align="left" valign="top" width="140">Enter Your Message:<br />
        <textarea cols="30" name="contact_text" class="textSize2 black" rows="5"></textarea> </td>
        </tr>
    <tr>
      <td>
      <input type="hidden" name="subject" value="Message from AndrewBrittonBooks.com">
      <input type="submit" value=" Submit " name="submit1"></td>
    </tr>
  </table>
  </form>

<% Else %>

	


	<%'mail using BRC client
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail")
	
	body = "<html><head><title></title></head><body bgcolor='#ffffff'>" & _
	"<b>First Name:</b><br>" & contact_firstname & "<br><br>" & _
	"<b>Last Name:</b><br>" & contact_lastname & "<br><br>" & _
	"<b>Email:</b> " & contact_email & "<br><br>" & _
	"<b>Additional Info:</b> " & contact_text & "<br><br>" & _
	"</body></html>"

	Mailer.FromAddress= contact_email
	Mailer.RemoteHost = "209.11.45.5"
	Mailer.AddRecipient "Andrew Britton", "Andrew@AndrewBrittonBooks.com"
	Mailer.Subject    = Request.Form ("subject")
	Mailer.contenttype = "text/html"
	Mailer.BodyText  = body
	if Mailer.SendMail then
	Response.Write "<font face='verdana, arial' size='2' color='#000000'><b>Thank you for your interest.</b> <br><br>Your submission has been sent.<br><br>"
	
	else
	Response.Write "Mail send failure. Error was " & Mailer.Response

	end if
	'end mail using BRC client %>


<% End If %>


<%
' The code below supplied by Mark Short 

' returns string the can be written where you would like the reCAPTCHA challenged placed on your page 
function recaptcha_challenge_writer(publickey) 
  recaptcha_challenge_writer = "<script type=""text/javascript"">" & _ 
  "var RecaptchaOptions = {" & _ 
  " theme : 'white'," & _ 
  " tabindex : 0" & _ 
  "};" & _ 
  "</script>" & _ 
  "<script type=""text/javascript"" src=""http://api.recaptcha.net/challenge?k=" & publickey & """></script>" & _ 
  "<noscript>" & _ 
  "<iframe src=""http://api.recaptcha.net/noscript?k=" & publickey & """ frameborder=""1""></iframe><br>" & _ 
  "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & _ 
  "<input type=""hidden"" name=""recaptcha_response_field"" value=""manual_challenge"">" & _ 
  "</noscript>" 
end function 

function recaptcha_confirm(privkey,rechallenge,reresponse) 
  ' Test the captcha field 
  Dim VarString 
  VarString = _ 
  "privatekey=" & privkey & _ 
  "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _ 
  "&challenge=" & rechallenge & _ 
  "&response=" & reresponse 
  Dim objXmlHttp 
  Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP") 
  objXmlHttp.open "POST", "http://api-verify.recaptcha.net/verify", False 
  objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
  objXmlHttp.send VarString 
  Dim ResponseString 
  ResponseString = split(objXmlHttp.responseText, vblf) 
  Set objXmlHttp = Nothing 
  if ResponseString(0) = "true" then 
    ' They answered correctly 
    recaptcha_confirm = "" 
  else 
    ' They answered incorrectly 
    recaptcha_confirm = ResponseString(1) 
  end if 
end function 
%>
<!-- END SEND/RESUBMIT -->






<% else %>
<!-- FIRST SUBMIT -->

<form action="contact.asp?submit=yes" method="post">
  <table cellspacing="2" cellpadding="0" border="0" width="490">
    <tr>
      <td align="left" width="140">First Name:<br /><input name="contact_firstname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_firstname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="left">Last Name:<br /><input name="contact_lastname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_lastname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="left">Email Address:<br /><input name="contact_email" style=" width: 200px;" class="textSize2 black" type="text" value="<%=contact_email%>" maxlength="150"></td>
    </tr>
    <tr>
      <td align="left" valign="top">Human Check:<br />
	  <%=recaptcha_challenge_writer(recaptcha_public_key)%></td>
      
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td align="left" valign="top" width="140">Enter Your Message:<br />
        <textarea cols="30" name="contact_text" class="textSize2 black" rows="5"></textarea> </td>
        </tr>
    <tr>
      <td>
      <input type="hidden" name="subject" value="Message from AndrewBrittonBooks.com">
      <input type="submit" value=" Submit " name="submit1"></td>
    </tr>
  </table>
  </form>


<%
' The code below supplied by Mark Short 

' returns string the can be written where you would like the reCAPTCHA challenged placed on your page 
function recaptcha_challenge_writer(publickey) 
  recaptcha_challenge_writer = "<script type=""text/javascript"">" & _ 
  "var RecaptchaOptions = {" & _ 
  " theme : 'white'," & _ 
  " tabindex : 0" & _ 
  "};" & _ 
  "</script>" & _ 
  "<script type=""text/javascript"" src=""http://api.recaptcha.net/challenge?k=" & publickey & """></script>" & _ 
  "<noscript>" & _ 
  "<iframe src=""http://api.recaptcha.net/noscript?k=" & publickey & """ frameborder=""1""></iframe><br>" & _ 
  "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & _ 
  "<input type=""hidden"" name=""recaptcha_response_field"" value=""manual_challenge"">" & _ 
  "</noscript>" 
end function 

function recaptcha_confirm(privkey,rechallenge,reresponse) 
  ' Test the captcha field 
  Dim VarString 
  VarString = _ 
  "privatekey=" & privkey & _ 
  "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _ 
  "&challenge=" & rechallenge & _ 
  "&response=" & reresponse 
  Dim objXmlHttp 
  Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP") 
  objXmlHttp.open "POST", "http://api-verify.recaptcha.net/verify", False 
  objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
  objXmlHttp.send VarString 
  Dim ResponseString 
  ResponseString = split(objXmlHttp.responseText, vblf) 
  Set objXmlHttp = Nothing 
  if ResponseString(0) = "true" then 
    ' They answered correctly 
    recaptcha_confirm = "" 
  else 
    ' They answered incorrectly 
    recaptcha_confirm = ResponseString(1) 
  end if 
end function 
%>

<!-- END FIRST SUBMIT -->
<% End if %>

<!-- ########################################## -->
<!-- ########################################## -->
<!-- #######  END CAPTCHA SUBMISSION FORM ##### -->
<!-- ########################################## -->
<!-- ########################################## -->
<!--
<FORM METHOD=POST ACTION="brittonmail.asp">
                      <B><FONT Class="text">Message:</FONT></B><BR>
                      <TEXTAREA name="message" rows="15" cols="28" wrap="virtual">Enter your message here and click 'Send'.</TEXTAREA>
                      <br>
                      <br>
                      <B><FONT Class="text">Your e-mail address:</FONT></B><BR>
                      <INPUT TYPE="text" NAME="email" size="30">
                      <input type="hidden" name="subject" value="From AndrewBrittonBooks.com - Contact Page">
                      <br>
                      <br>
                      <!--The names of the hidden fields are case sensitive. Only change their values 
          		  redirect specifies what page the user will be taken too after hitting send.
          		  recipient is where the email gets sent to.-->
                      
						<!-- INPUT TYPE="Submit" VALUE="Send">
                      <INPUT TYPE="Reset" VALUE="Clear">
                    </FORM>
			-->		