<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim ProblemsHTML, RaspunsuriHTML
 Dim cn
 
 set cn = Server.CreateObject("ADODB.Connection") 
 cn.Open Application("DSN")
 ProblemsHTML   = GetEnunturiProblemeHTML(Request.QueryString("PBIDs"), cn)
 RaspunsuriHTML = GetRaspunsuriProblemeHTML(Request.QueryString("PBIDs"), cn)  
 cn.Close
 set cn = nothing


 ' Intoarce textul HTML al unei pagini cu probleme
 Function GetEnunturiProblemeHTML(pbids, objCon)
  Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem IN (@1)"
  Dim PBDivs, contor, rs, re
  PBDivs = "<b>@1. @2</b><br>" & vbCrLf & "@3" & vbCrLf & "<br>" & vbCrLf & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", pbids))
  If not rs.EOF then
    contor = 1
    re = ""
    do until rs.EOF
      re = re & Replace(Replace(Replace(PBDivs, "@3", rs.Fields("textproblema").value), "@2", rs.Fields("numeproblema").value), "@1", CStr(contor))
      contor = contor + 1
      rs.movenext
    loop
  End if
  rs.Close
  set rs = nothing
  
  GetEnunturiProblemeHTML = re
 End Function

 ' Intoarce textul HTML al unei pagini cu raspunsurile de la probleme
 Function GetRaspunsuriProblemeHTML(pbids, objCon)
  Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem IN (@1)"
  Dim AnswsDIV, PBAnsw, contor, rs, re
  AnswsDIV = "<b>@1. @2</b><br>" & vbCrLf & "@3<br>" & vbCrLf & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", pbids))
  If not rs.EOF then
    contor = 1
    re = ""
    do until rs.EOF
      PBAnsw = GetPbAnswersStringForPrint(rs.Fields("id_problem").value, cn)
      re = re & Replace(Replace(Replace(AnswsDIV, "@1", contor), "@2", rs.Fields("numeproblema").value), "@3", PBAnsw) 
      contor = contor + 1
      rs.movenext
    loop
  End if
  rs.Close
  set rs = nothing
  
  GetRaspunsuriProblemeHTML = re
 End Function



' Intoarce sub forma de String o bucata HTML ce constituie zona cu raspunsuri
Function GetPbAnswersStringForPrint(pbid, objCon)
  Dim re, ServMachete, ServFinal
  Dim myCmd, rs
  Dim tipr, contor, optar, i

  re = ""
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbAnswers"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(pbid))
  If not rs.EOF then
    ServMachete = Array(_
        "<b>@1.</b><br>",_
        "<b>@1.</b><br>",_
        "<b>@1.</b> @3 - (@4)<br>",_
        "<b>@1.</b> @3<br>")
    contor = 1
    do until rs.EOF
     tipr = rs.Fields("tipraspuns").Value
     ServFinal = Replace(ServMachete(tipr-1),"@1", Chr(64+contor))
     select case tipr
       case 1,2 If not CBool(rs.Fields("responsecorrect").Value) then ServFinal = ""
       case 3   optar  = Split(rs.Fields("responsedetails").Value, Chr(3), -1, 1)
                ServFinal = Replace(Replace(ServFinal, "@4", Join(optar, ",")), "@3", optar(CInt(rs.Fields("responsecorrect").Value)))
       case 4   ServFinal = Replace(Replace(ServFinal,"@4", rs.Fields("responsedetails").Value),"@3", rs.Fields("responsecorrect").Value)
     end select
     re = re & ServFinal & vbCrLf
     contor = contor + 1
     rs.MoveNext
    loop
  GetPbAnswersStringForPrint = re  
  End If
  rs.Close
  set myCmd = nothing
  set rs = nothing
End Function
%>
<html>
<body>

<table id=tblEnunturi border=0 width=100%>
<tr><td align=left valign=middle bgcolor=#f0f0f0 height=30>
<font size=+1 color=#000064><b>Questions texts</b></font>
</td></tr>
<tr><td align=left valign=top style='padding-left:50px;'>
<%=ProblemsHTML%>
</td></tr>
</table>

<table id=tblRaspunsuri border=0 width=100%>
<tr><td align=left valign=middle bgcolor=#f0f0f0 height=30>
<font size=+1 color=#000064><b>Questions correct answers</b></font>
</td></tr>
<tr><td align=left valign=top style='padding-left:50px;'>
<%=RaspunsuriHTML%>
</td></tr>
</table>

<script language=vbscript>
 Sub Window_onload
  window.parent.HandleIframeLoading false
 End Sub 
</script>

</body>
</html>