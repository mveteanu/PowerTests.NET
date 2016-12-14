<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/HTMLPbControl.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim ProblemsHTML
 Dim cn
 
 set cn = Server.CreateObject("ADODB.Connection") 
 cn.Open Application("DSN")
 If Request.QueryString("PBIDs")<>"" then
   ProblemsHTML = GetCompleteProblemsHTML(Request.QueryString("PBIDs"), cn)
 Else
   ProblemsHTML = ""
 End If  
 cn.Close
 set cn = nothing

 ' Intoarce textul HTML COMPLET al unei pagini cu o problema
 Function GetCompleteProblemsHTML(pbids, objCon)
  Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem IN (@1)"
  Dim PBDivs, contor
  Dim rs, re

  PBDivs = "<div class='THTMLEditPageBorder' style='width:645px;'>" & vbCrLf &_
           "<div class='THTMLEditPage' style='width:645px;'>" & vbCrLf &_
           "<div style='display:block;padding:5px;'><b>@1. @2</b></div>" & vbCrLf &_
           "<div style='display:block;padding:5px;'>@3</div><br>" & vbCrLf &_
           "<div style='display:block;padding:5px;'><b>Correct answers:</b></div>" & vbCrLf &_
           "<div style='display:block;text-align:left;padding:5px;'>@4</div>" & vbCrLf &_
           "</div>" & vbCrLf &_
           "</div>" & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", pbids))
  If not rs.EOF then
    contor = 1
    re = ""
    do until rs.EOF
      PBAnsw = GetPbAnswersString(rs.Fields("id_problem").value, false, cn)
      re = re & Replace(Replace(Replace(Replace(PBDivs, "@4", PBAnsw), "@3", rs.Fields("textproblema").value), "@2", rs.Fields("numeproblema").value), "@1", CStr(contor))
      contor = contor + 1
      rs.movenext
    loop
  End if
  rs.Close
  set rs = nothing
  
  GetCompleteProblemsHTML = re
 End Function
%>
<html>
<head>
  <title>Preview questions</title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body>

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<fieldset id=GroupBox1 class=TGroupBox 
          style="width:698px;height:416px;"
          style="left:7px;top:8px;">
<legend>Questions</legend>
<div id="ViewPbDIV" style="position:absolute; left:8px; top:16px; width:680px; height:392px; background-color:threedshadow; border:inset thin; FONT-FAMILY:Times New Roman; FONT-Size: 12pt; overflow:auto;">
<%=ProblemsHTML%>
</div>
</fieldset>

<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:311px;top:432px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button1_onclick
  Window.close 
End Sub

</script>
</body>
</html>