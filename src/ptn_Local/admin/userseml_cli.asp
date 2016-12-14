<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
 Dim UsersNames
 Dim UsersEmails
 Dim SenderName
 Dim SenderEmail

 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSUsers = GetUsersByIDs(Request.QueryString("userlist"),cn)
 UsersNames = ""
 UsersEmails = ""
 do until RSUsers.EOF
   UsersNames  = UsersNames  & RSUsers.Fields("nume").Value &" "& RSUsers.Fields("prenume").Value 
   UsersEmails = UsersEmails &  RSUsers.Fields("email").Value 
   RSUsers.movenext
   if not RSUsers.EOF then 
     UsersNames = UsersNames & ", "
     UsersEmails = UsersEmails & ", "
   end if  
 loop
 RSUsers.Close
 set RSUsers=nothing

 Set RSUser = GetUserByID(Session("UserID"),cn)
 SenderName  = RSUser.Fields("nume").value & " " & RSUser.Fields("prenume").value
 SenderEmail = RSUser.Fields("email").value
 RSUser.Close
 Set RSUser = nothing
 
 cn.Close 
 set cn=nothing
%>
<html>
<head>
  <title>Send email</title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<form name="EmailForm" method=post target="FormReturn" action="userseml_ser.asp"> 
<fieldset class=TGroupBox style="width:481px;height:361px;"
          style="left:6px;top:8px;">
<legend>Email</legend>
<span id="Label1"
      class=TLabel style="width:28px;height:13px;"
      style="left:16px;top:32px;">
From:
</span>
<span id="Label2"
      class=TLabel style="width:34px;height:13px;"
      style="left:16px;top:64px;">
To:
</span>
<span id="Label3"
      class=TLabel style="width:39px;height:13px;"
      style="left:16px;top:96px;">
Subject:
</span>
<span id="Label4"
      class=TLabel style="width:100px;height:13px;"
      style="left:16px;top:120px;">
Message:
</span>
<input DISABLED id="Edit1" type=text value="<%=SenderName%>"
       class=TEdit style="width:401px;height:21px;"
       style="left:64px;top:24px;">
<input DISABLED id="Edit2" type=text value="<%=UsersNames%>"
       class=TEdit style="width:401px;height:21px;"
       style="left:64px;top:56px;">
<input name="msgsubject" id="Edit3" type=text value="Important message from PowerTests .NET"
       class=TEdit style="width:401px;height:21px;"
       style="left:64px;top:88px;">
<textarea name="msgbody" id="Memo1"
       class=TEdit style="width:449px;height:209px;"
       style="left:16px;top:136px;">
</textarea>
<input type=hidden name="msgfrom" value="<%=SenderEmail%>">
<input type=hidden name="msgto" value="<%=UsersEmails%>">
</fieldset>
<input id="Button1" type=button value="Cancel" title="Cancel sending email"
       class=TButton style="width:75px;height:25px;"
       style="left:253px;top:384px;">
<input id="Button2" type=submit value="Send" title="Send email"
       class=TButton style="width:75px;height:25px;"
       style="left:165px;top:384px;">
</form>
</div>

<div id="Form1Hidden" style="display:none;">
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  EmailForm.Memo1.focus  
End Sub


' Evenimentul apare la apasarea butonului Cancel
Sub Button1_onclick
  Window.close 
End Sub


' Evenimentul care apare la apasarea butonului Trimite (Submit)
Sub EmailForm_OnSubmit
  if EmailForm.Edit3.value="" or EmailForm.Memo1.value="" then
    msgbox "You need to enter email subject and message.", vbOkOnly+VbExclamation
    Window.event.returnValue = false
  else
    EmailForm.Button1.disabled=true
    EmailForm.Button2.disabled=true
  end if
End Sub
</script>

</body>
</html>
