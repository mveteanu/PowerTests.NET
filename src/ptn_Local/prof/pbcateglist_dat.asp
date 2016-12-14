<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetPbCategByCursID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|Category name|Questions"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_categpb").value & "|"
    .Write rs.Fields("numecateg").value & "|"
    .Write NullToZero(rs.Fields("pbincateg").value) & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>