<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetCursuriByStudent"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("UserID")))
 Response.Write "id|Course name|Professor"& vbCrLf
 do until rs.EOF
   with Response
    .Write IDCursWithLock(rs.Fields("id_curs").value, rs.Fields("studentblocat").value) & "|"
    .Write rs.Fields("numecurs").value & "|"
    .Write rs.Fields("numeprof").value & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
 
 Function IDCursWithLock(idcrs, lck)
  Dim re
  If lck then re = "L" & CStr(idcrs) else re = CStr(idcrs)
  IDCursWithLock = re
 End Function
%>