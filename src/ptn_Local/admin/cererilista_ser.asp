<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
 Const MsgVal_OK   = "Access granted to selected users. They should receive further information by email."
 Const MsgVal_Err1 = "Error: Some of the selected users were not added into DB due to some system errors."
 Const MsgVal_Err2 = "Error: System was not able to send email to some of the selected users."

 Const MsgDel_OK   = "The requests of selected users were denied. They should receive further information by email."
 Const MsgDel_Err1 = "Error: System was not able to send email to some of the selected users."
 Const MsgDel_Err2 = "Error: System was not able to remove requests of some selected users from DB although emails were sent to them."
 Const MsgDel_Err3 = "Error: System cannot remove selected requests from DB and cannot send emails."
 
 Dim Mesaj
 Dim Mesaj_Flags

 On Error Resume Next

 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")

 If Err.number<>0 then
   Mesaj = MsgGenErr
   Mesaj_Flags = vbOkOnly + vbCritical
 Else

	if Request.Form("SelectAction")="accept" then
	 Mesaj =  MsgVal_OK
	 Mesaj_Flags = vbOkOnly + vbInformation
	 select case ValidateUsersList(Request.Form("SelectList"), CLng(Session("UserID")), cn)
	   case 1 Mesaj       = MsgVal_Err1
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 2 Mesaj       = MsgVal_Err2
	          Mesaj_Flags = vbOkOnly + vbExclamation
	 end select
	else
	 Mesaj =  MsgDel_OK
	 Mesaj_Flags = vbOkOnly + vbInformation
	 select case DeleteUnvalidatedUsersList(Request.Form("SelectList"), CLng(Session("UserID")), cn)
	   case 1 Mesaj       = MsgDel_Err1
	          Mesaj_Flags = vbOkOnly + vbExclamation
	   case 2 Mesaj       = MsgDel_Err2
	          Mesaj_Flags = vbOkOnly + vbCritical
	   case 3 Mesaj       = MsgDel_Err3
	          Mesaj_Flags = vbOkOnly + vbCritical
	 end select
	end if

 End If
 
 If cn.State = adStateOpen then cn.Close
 Set cn=nothing
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox "<%=Mesaj%>",<%=Mesaj_Flags%>
</script>
</body>
