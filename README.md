<div align="center">

## Jmail email example in asp


</div>

### Description

The sole purpose of this article is to help those who could not figure out the right purpose of jmail object in asp....Just a simple example to send emails using jmail object...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[pompyk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pompyk.md)
**Level**          |Beginner
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pompyk-jmail-email-example-in-asp__4-7645/archive/master.zip)





### Source Code

<h1 align='center'>JMAIL:</h1>
<font color="green">' Create the JMail message Object</font><br>
set msg = Server.CreateOBject( "JMail.Message" )<br>
<br>
<font color="green">' Set logging to true to ease any potential debugging</font><br>
<font color="green">' And set silent to true as we wish to handle our errors ourself</font><br>
msg.Logging = true<br>
msg.silent = true<br><br>
<font color="green">' Most mailservers require a valid email address<br>
' for the sender</font><br>
<font color="green">'This email address I guess must be that which is provided when u booked
'your domain....</font><br>
<br>
msg.From = "changeme@change.com"<br>
msg.FromName = "SOMDUTT GROUND BOOKING"<br>
<br>
<font color="green">' Next we have to add some recipients.</font><br>
<font color="green">' The addRecipients method can be used multiple times.</font><br>
<font color="green">' Also note how we skip the name the second time, it</font><br>
<font color="green">' is as you see optional to provide a name.</font><br><br>
msg.AddRecipient "gangulysomdutt@yahoo.com", "SOMDUTT"<br>
<font color="green">'msg.AddRecipient "recipientelle@herDomain.com"</font><br>
<br>
<font color="green">' The subject of the message</font><br>
msg.Subject = "WHATEVER U LIKE"<br>
<font color="green">'don't use the same..u must make a html form...using it..</font><br>
<font color="green">'so that u can submit it...</font><br>
<font color="green">' Note the use of vbCrLf to add linebreaks to our email</font><br>
<br>
msg.Body ="Name: " & request.form("name") & vbcrlf & _<br>
"Address: " & request.form("address") & vbcrlf & _<br>
"City: " & request.form("city") & vbcrlf & _<br>
"Zip: " & request.form("zip") & vbcrlf & _<br>
"Phone: " & request.form("phone") & vbcrlf & _<br>
"Email: " & request.form("email") & vbcrlf & _<br>
"Fax: " & request.form("fax") & vbcrlf & _<br>
"Car: " & request.form("listcar") & vbcrlf & _<br>
"Bus: " & request.form("listbus") & vbcrlf & _<br>
"Utilisation Start Date: " & request.form("txtDate") & vbcrlf & _<br>
"Utilisation End Date: " & request.form("txtDate2") & vbcrlf & _<br>
"Payment Preference: " & request.form("select") & vbcrlf & _<br>
"Places to be visited: " & request.form("place") <br>
<font color="green">
' To send the message, you use the Send()<br>
' method, which takes one parameter that<br>
' should be your mailservers address<br>
'<br>
' To capture any errors which might occur,<br>
' we wrap the call in an IF statement<br>
'if not msg.Send("mail.myDomain.net" ) then<br>
</font>
<br>
if not msg.Send("202.71.139.79" ) then<br>
  Response.write "<pre>" & msg.log & "</pre>"<br>
else<br>
  Response.write "Message sent succesfully!"<br>
end if<br>
<font color="green">' bye bye...go u have done it!!..thx..<br><br>
'-------------------------------------------------<br>
'-----mail......ends...............................
<br></font>
<br>
This is just a little help to those who got entangled in JMAIL
<h3>gangulysomdutt@yahoo.com</h3>

