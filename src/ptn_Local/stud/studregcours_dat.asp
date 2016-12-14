<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetInscriereLaCursuri"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("UserID")))
 Response.Write "id|id_prof|Course name|Enrolled students|Enrollment requests|Maximum students|Enrollment policy"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_curs").value & "|"
    .Write rs.Fields("id_prof").value & "|"
    .Write rs.Fields("numecurs").value & "|"
    .Write rs.Fields("studentiinscrisi").value & "|"
    .Write rs.Fields("cereriinscriere").value & "|"
    .Write NrStudToString(rs.Fields("maxstudents").value, "") & "|"
    .Write PermisionToString(rs.Fields("permisiiacceptare").value) & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>