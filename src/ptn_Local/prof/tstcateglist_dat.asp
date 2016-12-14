<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetTstCategByCursID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|Category name|Tests|Visible"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_categtst").value & "|"
    .Write rs.Fields("numecateg").value & "|"
    .Write NullToZero(rs.Fields("tstincateg").value) & "|"
    .Write BooleanToAfirmColor(rs.Fields("categtstpublica").value, "Green", "Red") & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>