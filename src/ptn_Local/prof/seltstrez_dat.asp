<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetTestsSumarResByCursID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|Test name|Completed|Abandoned"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_test").value & "|"
    .Write rs.Fields("numetest").value & "|"
    .Write NullToZero(rs.Fields("NrRezolvari").value) & "|"
    .Write NullToZero(rs.Fields("NrNeterminate").value) & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>