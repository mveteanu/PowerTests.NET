<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
 Const Mesaj_DU  = "1. Personal information was saved succesfully."
 Const Mesaj_CP  = "2. Account password was changed."
 Const Mesaj_UP  = "2. Account password remained the same."
 Const Mesaj_WP  = "2. Account password was not changed (old password is incorrect)."
 Const Mesaj_PE  = "2. Error changing password."
 Const Mesaj_Err = "Error saving data."
 
 Dim UsersVals(5)
 Dim oldpass, newpass
 UsersVals(0) = Request.Form("nume").Item
 UsersVals(1) = Request.Form("prenume").Item
 UsersVals(2) = Request.Form("email").Item
 UsersVals(3) = Request.Form("telefon").Item
 UsersVals(4) = Request.Form("EditLang").Item
 oldpass = Request.Form("oldpass")
 newpass = Request.Form("newpass")
 
 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 if UpdateUser(Session("UserID"), UsersVals, cn) then
   Mesaj = Mesaj_DU
   Mesaj_Flags = vbOkOnly + vbInformation
   if oldpass<>"" then
     select case ChangeUserPass2(Session("UserID"), NewPass, OldPass, cn)
       case 0   Mesaj = Mesaj & """&vbCrLf&""" & Mesaj_CP
       case 1,2 Mesaj = Mesaj & """&vbCrLf&""" & Mesaj_PE
                Mesaj_Flags = vbOkOnly + vbCritical
       case 3   Mesaj = Mesaj & """&vbCrLf&""" & Mesaj_WP
                Mesaj_Flags = vbOkOnly + vbInformation
     end select
   else
     Mesaj = Mesaj & """&vbCrLf&""" & Mesaj_UP
   end if
 else
   Mesaj = Mesaj_Err
   Mesaj_Flags = vbOkOnly + vbCritical
 end if
 
 cn.Close 
 set cn=nothing
%>
<body>
<script language=vbscript>
  window.parent.ActivateButtons true
  msgbox "<%=Mesaj%>",<%=Mesaj_Flags%>
</script>
</body>
