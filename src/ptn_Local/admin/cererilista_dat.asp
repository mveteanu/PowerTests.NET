<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/Utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 tipuser=Request.QueryString("tipuser")
 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 set rsusers=GetUnvalidatedUsers(tipuser,cn)
 with Response
  .Write "id|last name|first name|email|phone|login|Sign-up date"& vbCrLf
  do until rsusers.eof
    eml = rsusers.fields("email").value
    .Write rsusers.fields("id_user").value & "|"
    .Write rsusers.fields("nume").value & "|"
    .Write rsusers.fields("prenume").value & "|"
    .Write "<a href='mailto:"& eml &"'>" & eml & "</a>|"
    .Write VidToSpace(rsusers.fields("telefon").value) & "|"
    .Write rsusers.fields("login").value & "|"
    .Write DToSR(rsusers.fields("datainscriere").value, "DD/MM/YYYY") & vbCrLf
    rsusers.movenext
  loop
 end with
 cn.Close 
 set cn=nothing
%>