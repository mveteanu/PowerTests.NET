<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetCursuriByProf"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("UserID")))
 Response.Write "id|Course name|Students|Maximum students|Enrollment policy|Public|Requests"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_curs").value & "|"
    .Write rs.Fields("numecurs").value & "|"
    .Write rs.Fields("studentiinscrisi").value & "|"
    .Write NrStudToString(rs.Fields("maxstudents").value, "") & "|"
    .Write PermisionToString(rs.Fields("permisiiacceptare").value) & "|"
    .Write BooleanToAfirmColor(rs.Fields("curspublic").value,"Green","Red") & "|"
    .Write IntToColor(rs.Fields("cereriinscriere").value, 0, "<>", "red", "") & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>