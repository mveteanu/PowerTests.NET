<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/Utils.asp" -->
<%
 Response.Buffer  = True
 Response.Expires = -1

 tipuser=Request.QueryString("tipuser")
 set cn=Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 set rsusers=GetValidatedUsers(tipuser,cn)
 with Response
  select case UCase(tipuser)
    case "A" .Write "id|Last name|First name|Email|Phone|Login|Account Date|Locked"& vbCrLf
    case "P" .Write "id|Last name|First name|Email|Phone|Login|Account Date|Courses|Students|Locked"& vbCrLf
    case "S" .Write "id|Last name|First name|Email|Phone|Login|Account Date|Courses|Locked"& vbCrLf
  end select
  do until rsusers.eof
    eml = rsusers.fields("email").value
    .Write rsusers.fields("id_user").value & "|"
    .Write rsusers.fields("nume").value & "|"
    .Write rsusers.fields("prenume").value & "|"
    .Write "<a href='mailto:"& eml &"'>" & eml & "</a> |"
    .Write VidToSpace(rsusers.fields("telefon").value) & "|"
    .Write rsusers.fields("login").value & "|"
    .Write DToSR(rsusers.fields("datavalidare").value, "DD/MM/YYYY") & "|"
    select case UCase(tipuser)
      case "P" .Write NullToZero(rsusers.fields("numarcursuri").value) & "|"
               .Write NullToZero(rsusers.fields("numarstudenti").value) & "|"
      case "S" .Write NullToZero(rsusers.fields("cursuriinscris").value) & "|"
    end select
    .Write BooleanToAfirmColor(rsusers.fields("locked").value, "red", "green") & vbCrLf
    rsusers.movenext
  loop
 end with
 cn.Close 
 set cn=nothing
%>

