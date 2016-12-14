<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/tests.asp" -->
<!-- #include file="../_serverscripts/HTMLPbControl.asp" -->
<%
 Response.Buffer  = True
 Response.Expires = -1

 Dim cn
 Dim qs, ProblemsHTML, PbArr, TInfo, FisaRezultateID
 
 Class TstInfo
  Public TestID
  Public Name
  Public NrProb
  Public Time
  Public NrSustinere
  Public MaxSustineri
  Public Comments
  Public StartDate
 End Class

 set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")

 If LCase(Request.QueryString("action")) = "dellastfisa" then
  cn.Execute "DELETE FROM TBStudentsResults WHERE id_fisarezultate=" & Request.Form("FisaRezultateID")
  With Response
   .Write "<script language=vbscript>"
   .Write "Window.Parent.Close"
   .Write "</script>"
   .End
  End With
 End If

 set TInfo = New TstInfo
 set tst1 = New PTNTestGenerator
 tst1.LoadTest Request.QueryString("TstID"), cn
 PbArr = tst1.GetGeneratedTest
 qs = Join(PbArr,",")
 
 TInfo.TestID       = tst1.TestDefinition.TestID
 TInfo.Name         = tst1.TestDefinition.Name
 TInfo.NrProb       = UBound(PbArr)+1
 TInfo.Time         = tst1.TestDefinition.Time
 TInfo.NrSustinere  = GetTestNrSustineriByUser(tst1.TestDefinition.TestID, Session("UserID"), cn) + 1
 TInfo.MaxSustineri = tst1.TestDefinition.MaxSustineri
 TInfo.Comments     = tst1.TestDefinition.Comments
 TInfo.StartDate    = tst1.GenerationDate
 set tst1 = nothing
 If qs<>"" then ProblemsHTML = GetCompleteProblemsHTML(qs, cn) else ProblemsHTML = ""
 FisaRezultateID = InsertFisaRezultate(Session("CursID"), Session("UserID"), TInfo, cn)
 cn.Close
 set cn = nothing

 ' Seteaza timpul de expirare a sesiunii
 If TInfo.Time = 0 then Session.Timeout  = 180 else  Session.Timeout  = TInfo.Time + 10
    

Function InsertFisaRezultate(cursid, userid, teinfo, objCon)
  Dim rs, re
  
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open "TBStudentsResults", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
  rs.AddNew
  re = rs.Fields("id_fisarezultate").Value
  rs.Fields("id_curs").Value       = CLng(cursid)
  rs.Fields("id_user").Value       = CLng(userid)
  rs.Fields("id_test").Value       = teinfo.TestID
  rs.Fields("datastarttest").Value = Now()
  rs.Fields("numetest").Value      = teinfo.Name
  rs.Fields("timp").Value          = teinfo.Time
  rs.Update 
  Set rs = nothing
  InsertFisaRezultate = re
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
           "<div style='display:block;padding:5px;'><b>Answers:</b></div>" & vbCrLf &_
           "<div style='display:block;text-align:left;padding:5px;'>@4</div>" & vbCrLf &_
           "</div>" & vbCrLf &_
           "</div>" & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", pbids))
  If not rs.EOF then
    contor = 1
    re = ""
    do until rs.EOF
      PBAnsw = GetStudPbAnswersString(rs.Fields("id_problem").value, cn)
      re = re & Replace(Replace(Replace(Replace(PBDivs, "@4", PBAnsw), "@3", rs.Fields("textproblema").value), "@2", rs.Fields("numeproblema").value), "@1", CStr(contor))
      contor = contor + 1
      rs.movenext
    loop
  End if
  rs.Close
  set rs = nothing
  
  GetCompleteProblemsHTML = re
 End Function

' Intoarce sub forma de String o bucata HTML ce constituie zona cu raspunsuri
' asa cum a fost completata de utilizator la momentul salvarii problemei
' Daca nu se gaseste problema se intoarce sirul vid
Function GetStudPbAnswersString(pbid, objCon)
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
    OptMacheta  = "<option value='@1'>@1</option>"
    ServMachete = Array(_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  id='answ@idpb_@idr' name='answgrp@idpb' type=radio class='TAnswerButton' title='Check correct answer'></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  id='answ@idpb_@idr' type=checkbox class='TAnswerButton' title='Check correct answers'></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><select id='answ@idpb_@idr' rows=1 class='TAnswerComboBox' title='Select correct answer'>@3</select></span>",_
        "<span unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input  id='answ@idpb_@idr' type=edit value='' class='TAnswerEdit' title='Type correct answer'></span>")
    contor = 1
    do until rs.EOF
     tipr = rs.Fields("tipraspuns").Value
     ServFinal = Replace(ServMachete(tipr-1), "@1", Chr(64+contor))
     ServFinal = Replace(Replace(ServFinal, "@idr", contor), "@idpb", pbid)
     If tipr = 3 then
       optar = Split(rs.Fields("responsedetails").Value, Chr(3), -1, 1)
       optstr = "<option value='' SELECTED></option>"
       For i = 0 to UBound(optar)
        optstr  = optstr & Replace(OptMacheta, "@1", optar(i))
       Next
       ServFinal = Replace(ServFinal,"@3", optstr)
     End If
     re = re & ServFinal & vbCrLf
     contor = contor + 1
     rs.MoveNext
    loop
    re = vbCrLf & "<input type=hidden id='pbresp' value='"& pbid &"|"& contor-1 &"'>" & vbCrLf & re
  GetStudPbAnswersString = re  
  End If
  rs.Close
  set myCmd = nothing
  set rs = nothing
End Function
%>
<html>
<head>
  <title>Test system</title>
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

<div id="Form0" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:32px;">
Test name:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:72px;">
Questions:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:112px;">
Previous solvings:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:152px;">
Maximum solvings:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:192px;">
Permited time:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:112px;top:232px;">
Comments:
</span>
<input READONLY type=text class=TEdit style="left:216px;top:32px;"
       style="width:409px;height:21px;background-color:buttonface;" value="<%=TInfo.Name%>">
<input READONLY type=text class=TEdit style="left:216px;top:72px;"
       style="width:409px;height:21px;background-color:buttonface;" value="<%=TInfo.NrProb%>">
<input READONLY type=text class=TEdit style="left:216px;top:112px;"
       style="width:409px;height:21px;background-color:buttonface;" value="<%=TInfo.NrSustinere-1%>">
<input READONLY type=text class=TEdit style="left:216px;top:152px;"
       style="width:409px;height:21px;background-color:buttonface;" value="<%=CompareToString2(TInfo.MaxSustineri, 0, "Unlimited")%>">
<input READONLY type=text class=TEdit style="left:216px;top:192px;"
       style="width:409px;height:21px;background-color:buttonface;" value="<%=CompareToString2(TInfo.Time, 0, "Unlimited")%>">
<textarea READONLY rows=2 class=TEdit style="left:216px;top:232px;"
       style="width:409px;height:185px;overflow:auto;background-color:buttonface;">
<%=TInfo.Comments%>
</textarea>       
<input id="ButtonStartTest" type=button value="Take test" title="Take this test"
       class=TButton style="width:100px;height:25px;"
       style="left:250px;top:460px;">
<input id="ButtonRenuntaTest" type=button value="Cancel" title="Cancel"
       class=TButton style="width:100px;height:25px;"
       style="left:370px;top:460px;">
</div>


<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<fieldset class=TGroupBox 
          style="width:698px;height:73px;"
          style="left:7px;top:8px;">
<legend>Test info</legend>
<span class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:24px;">
Name: <%=TInfo.Name%>
</span>
<span class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:48px;">
Questions: <%=TInfo.NrProb%>
</span>
<span class=TLabel style="width:142px;height:13px;"
      style="left:530px;top:24px;">
Solving no: <%=TInfo.NrSustinere & "/" & CompareToString2(TInfo.MaxSustineri, 0, "Unlimited")%>
</span>
<span id="TimeLabel" class=TLabel style="width:142px;height:13px;"
      style="left:530px;top:48px;">
Time: <%=0 & "/" & CompareToString2(TInfo.Time, 0, "Unlimited")%>
</span>
</fieldset>


<fieldset class=TGroupBox 
          style="width:698px;height:401px;"
          style="left:7px;top:83px;">
<legend>Questions</legend>
<div id="ViewPbDIV" style="position:absolute; left:8px; top:16px; width:680px; height:377px; background-color:threedshadow; border:inset thin; FONT-FAMILY:Times New Roman; FONT-Size: 12pt; overflow:auto;">
<%=ProblemsHTML%>
</div>
</fieldset>
<input id="ButtonAbandon" type=button value="Abandon" title="Cancel test"
       class=TButton style="width:210px;height:25px;"
       style="left:160px;top:496px;">
<input id="ButtonStopTest" type=button value="I'm done... show results" title="Stop test and display results"
       class=TButton style="width:210px;height:25px;"
       style="left:385px;top:496px;">
</div>

<form name="FormularH" method="post" action="studtstview_ser.asp" target="FormReturn">
<input type=hidden id="SelectAction"   name="SelectAction">
<input type=hidden id="SelectPBAnsw"   name="SelectPBAnsw">
<input type=hidden id="SelectTimeUsed" name="SelectTimeUsed">
<input type=hidden id="FisaRezultateID" name="FisaRezultateID" value=<%=CStr(FisaRezultateID)%>>
<input type=hidden id="TestName" name="TestName" value="<%=TInfo.Name%>">
<input type=hidden id="TestTime" name="TestTime" value="<%=TInfo.Time%>">
<input type=hidden id="TestStartDate" name="TestStartDate" value="<%=CStr(TInfo.StartDate)%>">
</form>

<div id="Form1Hidden" style="visibility:hidden;">
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% height=100% scrolling=no>
</IFRAME>
</div>

<div style="display:none;">
<IFRAME ID=FormDelReturn Name=FormDelReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
Dim NormalClose
Dim MaxTime, ElapsedTime, TimerID

' Evenimentul apare la incarcarea documentului
Sub window_onload
  NormalClose = false
  ElapsedTime = 0
  MaxTime = "<%=CompareToString2(TInfo.Time, 0, "Unlimited")%>"
  Form0.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Apare daca utilizatorul incearca sa inchida manual fereastra
Sub window_onbeforeunload
 If not NormalClose then window.event.returnValue = "Are you sure you want to cancel test?"
End Sub

' Evenimentul apare la incarcare in IFRAME-ul ascuns a noii pagini
' ce contine rezultatele testului
Sub FormReturn_onload
 NormalClose = true
 Form1Hidden.style.visibility = "visible"
 WaitforForm.style.visibility = "hidden"
End Sub

' Intoarce intr-o forma impachetata raspunsurile introduse de
' utilizator la o anumita problema
' rasppb = idpb #1 r1 #2 r2 #2 r3 #2 ...
Function GetPackedPbAnsw(pbid,nrrasp)
 Dim i, re
 re = CStr(pbid) & Chr(1)
 For i = 1 to nrrasp
  Select case LCase(ViewPbDIV.All("answ"&pbid&"_"&i).Type)
   case "text"              re = re & ViewPbDIV.All("answ"&pbid&"_"&i).Value & Chr(2)
   case "select-one"        re = re & CStr(ViewPbDIV.All("answ"&pbid&"_"&i).SelectedIndex - 1) & Chr(2)
   case "radio", "checkbox" re = re & CStr(CInt(ViewPbDIV.All("answ"&pbid&"_"&i).Checked)) & Chr(2)
  End Select
 Next
 re = Left(re, Len(re)-Len(Chr(2)))
 GetPackedPbAnsw = re
End Function

' Intoarce intr-o forma impachetata raspunsurile introduse de 
' utilizator la problemele din test
' rasptst = rasppb #0 rasppb #0 rasppb #0 rasppb ...
Function GetPackedUserAnswers
 Dim re
 Dim i, pbminfo
 
 re = ""
 <%If TInfo.NrProb-1 > 0 then%>
 For i=0 to <%=TInfo.NrProb-1%>
  pbminfo = Split(ViewPbDIV.All("pbresp")(i).Value, "|",-1,1)
  re = re & GetPackedPbAnsw(pbminfo(0), pbminfo(1)) & Chr(3)
 Next
 <%Else%>
  pbminfo = Split(ViewPbDIV.All("pbresp").Value, "|",-1,1)
  re = re & GetPackedPbAnsw(pbminfo(0), pbminfo(1)) & Chr(3)
 <%End If%>
 If re<>"" then re = Left(re,Len(re)-Len(Chr(3)))
 GetPackedUserAnswers = re
End Function

' Trateaza evenimentul care apare la apasarea butonului de renuntare la rezolvarea testului
Sub ButtonRenuntaTest_onclick
 NormalClose = true
 ButtonStartTest.disabled = true
 ButtonRenuntaTest.disabled = true
 FormularH.action = "studtstview.asp?action=dellastfisa"
 FormularH.target = "FormDelReturn"
 FormularH.submit 
End Sub

' Trateaza evenimentul care apare la apasarea butonului de rezolvare test
Sub ButtonStartTest_onclick
<%If ProblemsHTML<>"" then%>
 Form1.style.visibility = "visible"
 Form0.style.visibility = "hidden"
 TimerID = Window.setInterval("HandleTimerEvents", 60000)
<%Else%>
 msgbox "Cannot generate any test according to test definition.", vbOkOnly+vbExclamation
<%End If%>
End Sub

' Este executata automat la fiecare minut
Sub HandleTimerEvents
 ElapsedTime = ElapsedTime + 1
 TimeLabel.innerHTML = "Timp rezolvare: " & CStr(ElapsedTime) & "/" & CStr(MaxTime)
 If IsNumeric(MaxTime) then
  If ElapsedTime >= CInt(MaxTime) then 
    MsgBox "Time out!",vbOkOnly+vbInformation,"Warning"
    DoTestStopped
  End If  
 End If 
End Sub

' Executa operatiile necesare la terminarea testului, si anume 
' trimiterea pe server a raspunsurilor utilizatorilor in vederea
' validarii lor
Sub DoTestStopped
 Window.clearInterval TimerID
 With FormularH
    .SelectAction.value   = "StoreAnswers"
	.SelectPBAnsw.value   = GetPackedUserAnswers
	.SelectTimeUsed.value = CStr(ElapsedTime)
    WaitforForm.style.visibility = "visible"
    Form1.style.visibility = "hidden"
    .Submit 
 End With
End Sub

' Trateaza evenimentul aparut la apasarea butonului Abandoneaza test
Sub ButtonAbandon_onclick
 If MsgBox("Are you sure you want to abandon this test?" & vbCrLf& vbCrLf & "Abandoning a test will determine:" & vbCrLf & "1. storing the abandon in your profile;" & vbCrLf & "2. increment the number of times the test was taken.",vbYesNo+vbQuestion,"Confirm")=vbNo then Exit Sub
 msgbox "Test was abandoned!", vbOkOnly, "Info"
 NormalClose = true
 Window.close 
End Sub

' Trateaza evenimentul aparut la apasarea butonului Stop test
Sub ButtonStopTest_onclick
 Const StopMsg1 = "You're testing time was not elapsed yet. Are you sure you want to stop now?"
 Const StopMsg2 = "Are you sure you want to stop now?"
 Dim StopMsg

 If IsNumeric(MaxTime) then StopMsg = StopMsg1 else StopMsg = StopMsg2

 If MsgBox(StopMsg,vbYesNo+vbQuestion,"Confirm")=vbNo then Exit Sub
 DoTestStopped
End Sub
</script>

</body>
</html>