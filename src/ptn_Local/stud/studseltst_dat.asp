<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetTstInfoForStud"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,Array(CLng(Session("CursID")),CLng(Session("UserID"))))
 Response.Write "id|Test name|Time|Solvings|Maximum solvings"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_test").value & "|"
    .Write rs.Fields("numetest").value & "|"
    .Write CompareToString2(rs.Fields("timp").value, 0, "Unlimited") & "|"
    .Write NullToZero(rs.Fields("sustineri").value) & "|"
    .Write CompareToString2(rs.Fields("nrmaxsustineri").value, 0, "Unlimited") & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
%>