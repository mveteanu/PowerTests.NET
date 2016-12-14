<%@ Language=VBScript %>
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetCursuriByStudent"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Request.QueryString("iduser")))
 Response.Write "id|Course|Professor"& vbCrLf
 do until rs.EOF
   Response.Write "0|" & rs.Fields("numecurs").Value & "|" & rs.Fields("numeprof").Value & vbCrLf
   rs.MoveNext
 loop
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>