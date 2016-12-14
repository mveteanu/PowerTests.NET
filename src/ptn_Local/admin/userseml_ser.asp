<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/emails.asp" -->
<%
  Const MSG_OK  = "Email was sent successfully"
  Const MSG_Err = "Some error occured while sending the email."

  On Error Resume Next

  EmailAddrs = Split(Request.Form("msgto"), ",", -1, 1)
  for each adr in EmailAddrs
    SendEmail adr, Request.Form("msgfrom"), Request.Form("msgsubject"), Request.Form("msgbody")
  next
  
  If Err.number<>0 then
    Mesaj  = MSG_Err
    MesajF = vbOkOnly+vbCritical
  Else
    Mesaj  = MSG_OK
    MesajF = vbOkOnly+vbInformation
  End If
%>

<body>
<script language=vbscript>
 msgbox "<%=Mesaj%>",<%=MesajF%>
 window.parent.close
</script>
</body>