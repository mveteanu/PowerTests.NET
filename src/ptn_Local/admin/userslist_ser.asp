<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
 Const MsgGenErr    = "DB connection error."
 Const MsgGenErr2   = "Erro saving data into DB."
 Const MsgPassErr   = "Error changing password."
 Const MsgLockOK    = "Selected users were locked succesfully."
 Const MsgUnLockOK  = "Selected users were unlocked."
 Const MsgDeleteOK  = "Selected accounts were deleted."
 Const MsgUpdateOK  = "User data was saved succesfuly."
 Const MsgChPassOK  = "Password changed."
 Const MsgPassUnCh  = "Password remained unchanged." 
 
 Dim Action
 Dim UsersList
 Dim Mesaj
 Dim Mesaj_Flags

 Action    = Request.Form("SelectAction")
 UsersList = Request.Form("SelectList")
 UsersVals = Split(Request.Form("SelectValues"),"|",-1,1)

 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")

 If Err.number<>0 then
   Mesaj = MsgGenErr
   Mesaj_Flags = vbOkOnly + vbCritical
 Else

 Select Case Action
   case "lock"   if LockUsers(UsersList, true, cn) then
                   Mesaj = MsgLockOK
                   Mesaj_Flags = vbOkOnly + vbInformation
                 else
                   Mesaj = MsgGenErr2
                   Mesaj_Flags = vbOkOnly + vbCritical
                 end if  
   case "unlock" if LockUsers(UsersList, false, cn) then
                   Mesaj = MsgUnLockOK
                   Mesaj_Flags = vbOkOnly + vbInformation
                 else
                   Mesaj = MsgGenErr2
                   Mesaj_Flags = vbOkOnly + vbCritical
                 end if
   case "delete" if DeleteUsers(UsersList, cn) then
                   Mesaj = MsgDeleteOK
                   Mesaj_Flags = vbOkOnly + vbInformation
                 else
                   Mesaj = MsgGenErr2
                   Mesaj_Flags = vbOkOnly + vbCritical
                 end if
   case "edit"   if UpdateUser(CLng(UsersList), UsersVals, cn) then
                   Mesaj = MsgUpdateOK
                   Mesaj_Flags = vbOkOnly + vbInformation
                   if Len(UsersVals(5)) > 0 then
                     if ChangeUserPass1(CLng(UsersList), UsersVals(5), cn)=0 then
                       Mesaj = Mesaj & """&vbCrLf&""" & MsgChPassOK
                     else
                       Mesaj = Mesaj & """&vbCrLf&""" & MsgPassErr
                       Mesaj_Flags = vbOkOnly + vbCritical
                     end if  
                   else
                       Mesaj = Mesaj & """&vbCrLf&""" & MsgPassUnCh
                   end if
                 else
                   Mesaj = MsgGenErr2
                   Mesaj_Flags = vbOkOnly + vbCritical
                 end if
 End Select
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
