<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/tests.asp" -->
<!-- #include file="../_serverscripts/HTMLPbControl.asp" -->
<!-- #include file="../_serverscripts/TabControl.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 Dim cn
 Dim ProblemsHTML
 Dim UserAnswers, PBCSVList
 Dim PBResults, PointsObtained, MaxPoints, NotaObtinuta

 UserAnswers = Split(Request.Form("SelectPBAnsw"),Chr(3),-1,1)
 set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 PBCSVList = GetPbList(UserAnswers)

 Set PBResults  = GetTestResultsDict(PBCSVList, cn)
 PointsObtained = GetTestPointsObtained(PBResults)
 MaxPoints      = GetTestPointsMax(PBCSVList)
 NotaObtinuta   = (PointsObtained*10)/MaxPoints
 InsertResultsInDB Request.Form("FisaRezultateID"), MaxPoints, PointsObtained, cn

 ProblemsHTML  = GetCompleteProblemsHTML(PBCSVList, cn)
 cn.Close
 Set cn = nothing

' Insereaza sumarul rezultatelor testului in baza de date
Sub InsertResultsInDB(idfisa, pcttotal, pctobtin, objCon)
	objCon.StudResultsUpdate CLng(idfisa), CDbl(pcttotal), CDbl(pctobtin)
End Sub

' Compara doua siruri in vederea stabilirii egalitatii lor prin 
' potrivirea specificata de metoda metpotr
Function CompareStrings(sstu, scor, ignoreblk, casesens, metpotr)
 Dim re
 Dim sstuw, scorw
 
 re = false
 If ignoreblk then 
   sstuw = Trim(sstu): scorw = Trim(scor)
 Else
   sstuw = sstu: scorw = scor
 End If 
 If not casesens then sstuw = LCase(sstuw): scorw = LCase(scorw)
 
 If (sstu<>"") and (scor<>"") then 
	Select case metpotr
	  Case 0 If sstuw = scorw then re = true
	  Case 1 If sstuw = Left (scorw, Len(sstuw)) then re = true
	  Case 2 If sstuw = Right(scorw, Len(sstuw)) then re = true
	  Case 3 If InStr(1, scorw, sstuw, 1)>0 then re = true
	End Select
 End If

 CompareStrings = re
End Function

' Intoarce true daca un raspuns de tip edit ce are proprietatile
' propr este corect. answstud si answcorr sunt valorile ce se compara
Function IsEditAnswerCorrect(answstud, answcorr, propr)
 Dim re
 Dim AnswWorkStud
 Dim AnswPrecizie, AnswIgnoreNonNum, AnswNrStu, AnswNrCor
 Dim AnswMatchMet, AnswCaseSens, AnswIgnoreBlack
 
 On Error Resume Next
 re = false
 If ((propr and 1) = 0) then     ' Raspuns numeric
   AnswPrecizie      = (propr and 14)/2
   AnswIgnoreNonNum  = ((propr and 16) = 16)
   If AnswIgnoreNonNum then 
     AnswNrStu = CDbl(ExtractNumeric(answstud))
     AnswNrCor = CDbl(ExtractNumeric(answcorr))
   Else  
     AnswNrStu = CDbl(GetRegionalDouble(answstud))
     AnswNrCor = CDbl(GetRegionalDouble(answcorr))
   End If
   If AnswPrecizie <> 7 then
    AnswNrStu = Round(AnswNrStu, AnswPrecizie)
    AnswNrCor = Round(AnswNrCor, AnswPrecizie)
   End If
   If AnswNrStu = AnswNrCor then re = true
 Else                            ' Raspuns alfanumeric
   AnswMatchMet      = (propr and 14)/2
   AnswCaseSens      = ((propr and 16) = 16)
   AnswIgnoreBlack   = ((propr and 32) = 32)
   re = CompareStrings(answstud, answcorr, AnswIgnoreBlack, AnswCaseSens, AnswMatchMet)
 End If
 
 If Err.number<>0 then re = false: Err.Clear 

 IsEditAnswerCorrect = re
End Function


 ' Extrage din informatia primita prin formular un string in format
 ' CSV ce contine ID-urile problemelor din test
 Function GetPbList(userresp)
  Dim re
  Dim raspar2, ra
  
  re = ""
  For Each ra In userresp
   raspar2 = Split(ra,Chr(1),-1,1)
   re = re & raspar2(0) & ","
  Next
  If re<>"" then re = Left(re, Len(re)-Len(","))
  GetPbList = re
 End Function

 ' Obtine rezultatul dat de student la un anumit raspuns al unei
 ' probleme. Sunt folosite informatiile primite prin formular
 Function GetPbStudAnswValue(idp, nra, userresp)
  Dim re, ra, raspar1, raspar2
  
  For Each ra in userresp
   raspar1 = Split(ra, Chr(1), -1, 1)
   If CInt(raspar1(0)) = CInt(idp) then
     raspar2 = Split(raspar1(1), Chr(2), -1, 1)
     If raspar1(1)<>"" then re = raspar2(nra-1) else re = ""
     Exit For
   End If  
  Next
  GetPbStudAnswValue = re
 End Function


 ' Intoarce un calificativ in functie de punctele obtinute la
 ' o problema
 Function PbResultToCalificativ(optainedpoints, totalpoints)
  If optainedpoints = 0 then
    PbResultToCalificativ  = "<font color=red><b>Bad answer! Obtained points: "&Round(optainedpoints,2)&"</b></font>"
  ElseIf optainedpoints >= totalpoints then
    PbResultToCalificativ  = "<font color=red><b>Correct! Obtained points: "&Round(optainedpoints,2)&"</b></font>"
  Else
    PbResultToCalificativ  = "<font color=red><b>Partially correct! Obtained points: "&Round(optainedpoints,2)&"</b></font>"
  End If
 End Function


 ' Intoarce un obiect de tip Dictionary ce contine rezultatele
 ' obtinute la fiecare problema. Pe post de keye este ID-ul 
 ' problemei si pe post de continut se afla rezultatul obtinut
 Function GetTestResultsDict(pbids, objCon)
   Dim re, pbar, pba
   Set re = CreateObject("Scripting.Dictionary")
   pbar = Split(pbids, ",", -1, 1)
   For Each pba in pbar
    re.Add pba, ComputePbResult(pba, UserAnswers, objCon)
   Next
   Set GetTestResultsDict = re
 End Function

 ' Intoarce punctele totale ale testului
 ' Practic fiecare problema valoreaza un punct si deci punctele = nrpb
 Function GetTestPointsMax(pbids)
  Dim pbar
  pbar = Split(pbids, ",", -1, 1)
  GetTestPointsMax = UBound(pbar) + 1
 End Function
 
 ' Intoarce punctele obtinute
 Function GetTestPointsObtained(resultdict)
  Dim re
  re = 0
  For Each k in resultdict.Keys
   re = re + CDbl(resultdict.Item(k))
  Next
  GetTestPointsObtained = re
 End Function

 ' Intoarce nr. de puncte obtinute la o anumita problema
 Function ComputePbResult(pbid, userresp, objCon)
  Dim myCmd, rs, re
  Dim acceptpartial, tiprasp
  Dim answcorecte, answtotal
  
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbAnswers"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(pbid))
  re = 0
  If not rs.EOF then
    acceptpartial = rs.Fields("acceptaraspunspartial").value
    tiprasp       = rs.Fields("tipraspuns").value
	answcorecte   = 0
	answtotal     = 0
	Do until rs.EOF
	  answtotal   = answtotal + 1
	  Select Case tiprasp
	   case 1, 2 If CBool(rs.Fields("responsecorrect").value) = CBool(GetPbStudAnswValue(pbid, answtotal, userresp)) then answcorecte = answcorecte + 1
	   case 3    If CInt(rs.Fields("responsecorrect").value)  = CInt(GetPbStudAnswValue(pbid, answtotal, userresp)) then answcorecte = answcorecte + 1
	   case 4    If IsEditAnswerCorrect(rs.Fields("responsecorrect").value, GetPbStudAnswValue(pbid, answtotal, userresp), rs.Fields("responsedetails").value) then answcorecte = answcorecte + 1
	  End Select
	  rs.MoveNext
	loop
	If acceptpartial then 
	  re = answcorecte / answtotal 
	Else
	  If answcorecte = answtotal then re = 1 else re = 0
	End If
	If (((tiprasp = 1) or (tiprasp = 2)) and (re<1)) then re = 0
  End If
  rs.Close
  set myCmd = nothing
  set rs = nothing
  
  ComputePbResult = re
 End Function


 ' Intoarce textul HTML COMPLET al unei pagini cu o problema
 Function GetCompleteProblemsHTML(pbids, objCon)
  Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem IN (@1)"
  Dim PBDivs, contor
  Dim rs, re

  PBDivs = "<div class='THTMLEditPageBorder' style='width:645px;'>" & vbCrLf &_
           "<div class='THTMLEditPage' style='width:645px;'>" & vbCrLf &_
           "<div style='display:block;padding:5px;'><b>@1. @2</b></div>" & vbCrLf &_
           "<div style='display:block;padding:5px;'>@3</div><br>" & vbCrLf &_
           "<table bgcolor=#f5f5f5 width=100% border=0 style='border-top:1px solid black;'><tr>" & vbCrLf &_
           "<td colspan=2><div style='display:block;text-align:left;padding:5px;width:300px;'>@puncte</div></td></tr><tr>" & vbCrLf &_
           "<td><div style='display:block;text-align:left;padding:5px;width:300px;'><b>Student answers:</b><br><br>@5</div></td>" & vbCrLf &_
           "<td><div style='display:block;text-align:left;padding:5px;width:300px;'><b>Correct answers:</b><br><br>@4</div></td>" & vbCrLf &_
           "</tr></table>" & vbCrLf &_
           "</div>" & vbCrLf &_
           "</div>" & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", pbids))
  If not rs.EOF then
    contor = 1
    re = ""
    do until rs.EOF
      PBAnsw  = GetPbAnswersString(rs.Fields("id_problem").value, false, cn)
      PBAnsw2 = GetStudPbAnswersString2(rs.Fields("id_problem").value, UserAnswers, cn)
      re = re & Replace(Replace(Replace(Replace(Replace(Replace(PBDivs, "@puncte", PbResultToCalificativ(PBResults.Item(CStr(rs.Fields("id_problem").value)),1)), "@5", PBAnsw2), "@4", PBAnsw), "@3", rs.Fields("textproblema").value), "@2", rs.Fields("numeproblema").value), "@1", CStr(contor))
      contor = contor + 1
      rs.movenext
    loop
  End if
  rs.Close
  set rs = nothing
  
  GetCompleteProblemsHTML = re
 End Function


' Intoarce sub forma de String o bucata HTML ce constituie zona cu raspunsuri
' Daca nu se gaseste problema se intoarce sirul vid
Function GetStudPbAnswersString2(pbid, userresp, objCon)
  Dim re, ServMachete, OptMacheta, ServFinal
  Dim myCmd, rs
  Dim tipr, contor, optstr, optar, i

  re = ""
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbAnswers"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(pbid))
  If not rs.EOF then
    OptMacheta  = "<option value='@1'@2>@1</option>"
    ServMachete = Array(_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  type=radio class='TAnswerButton' @chk></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  type=checkbox class='TAnswerButton' @chk></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><select rows=1 class='TAnswerComboBox'>@3</select></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  readonly type=edit class='TAnswerEdit' value='@edval'></span>")
    contor = 1
    do until rs.EOF
     tipr = rs.Fields("tipraspuns").Value
     ServFinal = Replace(ServMachete(tipr-1), "@1", Chr(64+contor))
     ServFinal = Replace(Replace(ServFinal, "@idr", contor), "@idpb", pbid)
     Select case tipr
       case 1,2 If CBool(GetPbStudAnswValue(CLng(pbid), contor, UserAnswers)) then
                  ServFinal = Replace(ServFinal,"@chk", "CHECKED")
                Else
                  ServFinal = Replace(ServFinal,"@chk", "")
                End If
       case 3 optar = Split(rs.Fields("responsedetails").Value, Chr(3), -1, 1)
              optstr = "<option value=''></option>"
              For i = 0 to UBound(optar)
				optstr2 = Replace(OptMacheta, "@1", optar(i))
				If i = CInt(GetPbStudAnswValue(CLng(pbid), contor, UserAnswers)) then
				  optstr2 = Replace(optstr2, "@2", " SELECTED ")
				Else
				  optstr2 = Replace(optstr2, "@2", "")
				End If  
				optstr = optstr & optstr2
			  Next
			  ServFinal = Replace(ServFinal,"@3", optstr)
	   case 4 ServFinal = Replace(ServFinal,"@edval", GetPbStudAnswValue(CLng(pbid), contor, UserAnswers))
     End Select
     re = re & ServFinal & vbCrLf
     contor = contor + 1
     rs.MoveNext
    loop
  GetStudPbAnswersString2 = re  
  End If
  rs.Close
  set myCmd = nothing
  set rs = nothing
End Function
%>
<html>
<head>
  <title>Test results</title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
  <script language=javascript src="../_clientscripts/tabControlEvents.js"></script>
</head>
<body>

<div id="Form1" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<%OpenTabControl 8, 4, 708, 484, Array("Summary","Details"), 1, "PageControl1"%>

<%OpenTabContent%>
<span class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:8px;">
Test results summary:
</span>
<div style="position:absolute; left:8px; top:24px; width:680px; height:420px; background-color:threedshadow; border:inset thin; FONT-FAMILY:Times New Roman; FONT-Size: 12pt; overflow:auto;">
<div class='THTMLEditPageBorder' style='width:645px;'>
<div class='THTMLEditPage' style='width:645px;background-color:white;color:black;'>
<table border=0 width=95% align=center cellspacing=10>
<tr>
<td colspan=2>
<table border=0 width=100% align=center style='background-color:#f5f5f5;border-top:1px solid black; border-bottom:1px solid black;'>
<tr>
<td width=60%>
Test name: <%=Request.Form("TestName")%><br>
Starting date: <%=Request.Form("TestStartDate")%><br>
Used time: <%=Request.Form("SelectTimeUsed")%> minutes (Available time: <%=Request.Form("TestTime")%> minutes)<br>
</td>
<td>
Test points: <%=MaxPoints%><br>
Obtained points: <%=Round(PointsObtained,3)%><br>
Score: <%=Round(NotaObtinuta,2)%><br>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td width=50%>
<img src="studtstview_img.asp?imgid=2&v1=<%=Request.Form("TestTime")%>&v2=<%=Request.Form("SelectTimeUsed")%>">
</td>
<td width=50%>
<img src="studtstview_img.asp?imgid=1&v1=<%=MaxPoints%>&v2=<%=Round(PointsObtained,3)%>">
</td>
</tr>
<tr>
<td colspan=2 align=center>
<img src="studtstview_img.asp?imgid=3&packedpbs=<%=PackDictToString(PBResults,"=", ";")%>">
</td>
</tr>
</table>
</div>
</div>
</div>
<%CloseTabContent%>

<%OpenTabContent%>
<span class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:8px;">
Detailed results:
</span>
<div id="ViewPbDIV2" style="position:absolute; left:8px; top:24px; width:680px; height:420px; background-color:threedshadow; border:inset thin; FONT-FAMILY:Times New Roman; FONT-Size: 12pt; overflow:auto;">
<%=ProblemsHTML%>
</div>

<%CloseTabContent%>
<%CloseTabControl%>

<input id="ButtonCloseRezultate" type=button value="Close" title="Close form"
       class=TButton style="width:100px;height:25px;"
       style="left:310px;top:496px;">
</div>

<script language=vbscript>
 Sub ButtonCloseRezultate_onclick
  window.parent.close 
 End Sub
</script>

</body>
</html>