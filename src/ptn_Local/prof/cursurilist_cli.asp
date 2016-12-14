<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:640px;height:298px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:640px;height:298px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 11, 8, 233, Array("Course name", "Requests", "Students", "Maximum students", "Enrollment policy", "Public"), Array(130,70,80,120,140,70), 1, "cursurilist_dat.asp" , true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Add course" title="Add a new course"
       class=TButton style="width:85px;height:25px;"
       style="left:10px;top:256px;">
<input DISABLED id="Button2" type=button value="Delete course" title="Delete selected course and all associated data"
       class=TButton style="width:85px;height:25px;"
       style="left:115px;top:256px;">
<input DISABLED id="Button3" type=button value="Edit" title="Edit course properties"
       class=TButton style="width:85px;height:25px;"
       style="left:220px;top:256px;">
<input DISABLED id="Button4" type=button value="Info couse" title="View course summary"
       class=TButton style="width:85px;height:25px;"
       style="left:325px;top:256px;">
<input DISABLED id="Button5" type=button value="Go to class" title="Go to class!"
       class=TButton style="width:85px;height:25px;"
       style="left:430px;top:256px;">
<input id="Button6" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:535px;top:256px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="cursurilist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValues" name="SelectValues">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>


<script language=vbscript>
' Schimba starea activ/inactiv a butoanelor
Sub ActivateButtons(btnsstate)
 Dim i
 
 for i=1 to 6
   Form1.all("Button"&CStr(i)).disabled = not btnsstate
 next  
End Sub


' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub


' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  ActivateButtons true
End Sub


' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub



' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Intoarce sub forma de string ID-ul recordului selectat.
' Daca nu se selecteaza nici o inregistrare sau se selecteaza mai mult de una
' atunci se afiseaza un mesaj si se intoarce sirul vid.
Function GetSelectedRecord
  Dim RecList
  Dim RecArray
  
  RecList  = TableGetSelected(tblMyTable)
  RecArray = Split(RecList,",",-1,1)
  If (UBound(RecArray)-LBound(RecArray))<>0 then 
    MsgBox "You need to select a record first.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function


' Evenimentul apare la apasarea butonului Add curs
Sub Button1_OnClick
  DateCurs = ShowModalDialog("cursurilist_cliadd.asp",, "dialogWidth=348px;dialogHeight=400px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  if not IsArray(DateCurs) then Exit Sub
  
  FormularH.SelectAction.value = "add"
  FormularH.SelectValues.value = DateCurs(0) & "|" & DateCurs(1) & "|" & DateCurs(2) & "|" & DateCurs(3)
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul apare la apasarea butonului Delete curs
Sub Button2_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to delete selected course?",vbYesNo+vbQuestion,"Confirm") = vbNo then
    Exit Sub
  Else  
    if msgbox("Are you absolutely sure you want to delete selected course?",vbYesNo+vbQuestion,"Confirm again") = vbNo then Exit Sub
  End If     

  FormularH.SelectAction.value = "del"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul apare la apasarea butonului Edit curs
Sub Button3_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  DateCurs = ShowModalDialog("cursurilist_cliadd.asp?cursid=" & RecList ,, "dialogWidth=348px;dialogHeight=400px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  if not IsArray(DateCurs) then Exit Sub
  
  FormularH.SelectAction.value = "edit"
  FormularH.SelectList.value = RecList
  FormularH.SelectValues.value = DateCurs(0) & "|" & DateCurs(1) & "|" & DateCurs(2) & "|" & DateCurs(3)
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Info curs
Sub Button4_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  ShowModalDialog "cursinfo_cli.asp?cursid=" & RecList,, "dialogWidth=660px;dialogHeight=364px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub


' Evenimentul care apare la apasarea butonului Intra la curs
Sub Button5_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  window.parent.frames("Header").location.href = "headerprof.asp?CursID=" & RecList
  HideAllDivs
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button6_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
