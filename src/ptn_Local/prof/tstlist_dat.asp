<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetTstInfo"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|Test name|Time|Use quest|Not use quest|Random q.|Solvings|Max Solvings|Visible"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_test").value & "|"
    .Write rs.Fields("numetest").value & "|"
    .Write NumberToNelim(rs.Fields("timp").value) & "|"
    .Write NullToZero(rs.Fields("includepb").value) & "|"
    .Write NullToZero(rs.Fields("excludepb").value) & "|"
    .Write NullToZero(rs.Fields("maxrandom").value) & "|"
    .Write NullToZero(rs.Fields("NrSustineri").value) & "|"
    .Write NumberToNelim(NullToZero(rs.Fields("nrmaxsustineri").value)) & "|"
    .Write BooleanToAfirmColor(rs.Fields("testpublic").value,"Green","Red") & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
 
 
 Function NumberToNelim(nr)
  If CInt(nr) <> 0 then 
    NumberToNelim = CStr(nr)
  Else
    NumberToNelim = "Unlimited" 
  End If  
 End Function
%>