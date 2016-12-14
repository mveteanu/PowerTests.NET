<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
 Const MsgListaErr = "Error obtaining requests list."
 Const MsgGenErr   = "DB connection error"

 Const MsgVal_OK   = "Access granted to all users. They should receive further information by email."
 Const MsgVal_Err1 = "Error: Some of users were not validated into DB due to some system errors."
 Const MsgVal_Err2 = "Error: System was not able to send email to some users."

 Const MsgDel_OK   = "All requests were denied. They should receive further information by email."
 Const MsgDel_Err1 = "Error: System was not able to remove requests of some users from DB although emails were sent to them."
 Const MsgDel_Err2 = "Error: System was not able to send email to some users."
 Const MsgDel_Err3 = "Error: System cannot remove requests from DB and cannot send emails to users."
 
 Dim Mesaj
 Dim Mesaj_Flags

 On Error Resume Next

 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")

 If Err.number<>0 then
   Mesaj = MsgGenErr
   Mesaj_Flags = vbOkOnly + vbCritical
 Else

	if Request.Form("SelectAction")="acceptall" then
	 Mesaj =  MsgVal_OK
	 Mesaj_Flags = vbOkOnly + vbInformation
	 select case ValidateAllUsers(CLng(Session("UserID")), cn)
	   case 1 Mesaj       = MsgListaErr
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 2 Mesaj       = MsgVal_Err1
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 4 Mesaj       = MsgVal_Err2
	          Mesaj_Flags = vbOkOnly + vbExclamation
	 end select
	else
	 Mesaj =  MsgDel_OK
	 Mesaj_Flags = vbOkOnly + vbInformation
	 select case DeleteAllUnvalidatedUsers(CLng(Session("UserID")), cn)
	   case 1 Mesaj       = MsgListaErr
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 2 Mesaj       = MsgDel_Err1
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 4 Mesaj       = MsgDel_Err2
	          Mesaj_Flags = vbOkOnly + vbExclamation
	   case 6 Mesaj       = MsgDel_Err3
	          Mesaj_Flags = vbOkOnly + vbCritical
	 end select
	end if

 End If
 
 If cn.State = adStateOpen then cn.Close
 Set cn=nothing
%>
<body>
<script language=vbscript>
  msgbox "<%=Mesaj%>",<%=Mesaj_Flags%>
  window.parent.HideAllDivs
</script>
</body>
