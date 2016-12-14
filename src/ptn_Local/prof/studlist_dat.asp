<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetStudValByCurs"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|IDStud|Last name|First name|Email|Phone|Login|No of testings|Locked"& vbCrLf
 do until rs.EOF
   with Response
    eml = rs.fields("email").value
    .Write rs.Fields("id_studcurs").value & "|"
    .Write rs.Fields("id_user").value & "|" 
    .Write rs.Fields("nume").value & "|"
    .Write rs.Fields("prenume").value & "|"
    .Write "<a href='mailto:"& eml &"'>" & eml & "</a>|"
    .Write rs.Fields("telefon").value & "|"
    .Write rs.Fields("login").value & "|"
    .Write NullToZero(rs.Fields("TesteRezolvate").value) & "|"
    .Write BooleanToAfirmColor(rs.fields("locked").value, "Red", "Green") & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>