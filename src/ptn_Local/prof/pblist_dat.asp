<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetPbInfo"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(Session("CursID")))
 Response.Write "id|Question name|Answer type|Partial answer|Answers|Images|Categories|Tests"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_problem").value & "|"
    .Write rs.Fields("numeproblema").value & "|"
    .Write RaspunsIDToStr(rs.Fields("tipraspuns").value) & "|"
    .Write BooleanToAfirm(rs.Fields("acceptaraspunspartial").value) & "|"
    .Write NullToZero(rs.Fields("nransw").value) & "|"
    .Write NullToZero(rs.Fields("nrimag").value) & "|"
    .Write NullToZero(rs.Fields("nrcateg").value) & "|"
    .Write NullToZero(rs.Fields("nrtests").value) & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing


 ' Returneaza numele tipului raspunsului in fn. de codul sau
 Function RaspunsIDToStr(tiprasp)
  Dim re
  
  Select case tiprasp
    case 1 re = "Radio"
    case 2 re = "Check"
    case 3 re = "Combo"
    case 4 re = "Edit"
  End Select
  
  RaspunsIDToStr = re  
 End Function
%>