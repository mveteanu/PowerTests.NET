<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/TableControl.asp" -->
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:473px;height:275px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;" unselectable="on"
     class=TForm style="width:473px;height:275px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 8, 8, 209, Array("Course name", "Professor"), Array(249,200), 1, "studentercours_dat.asp" , true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Un-enroll" title="Un-enroll from selected course"
       class=TButton style="width:85px;height:25px;"
       style="left:8px;top:232px;">
<input DISABLED id="Button2" type=button value="Info" title="More information about selected course"
       class=TButton style="width:85px;height:25px;"
       style="left:129px;top:232px;">
<input DISABLED id="Button3" type=button value="Go to class" title="Go to class"
       class=TButton style="width:85px;height:25px;"
       style="left:251px;top:232px;">
<input id="Button4" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:372px;top:232px;">
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="studregcours_ser.asp" target="FormReturn">
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
 
 for i=1 to 3
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
    MsgBox "You need to select a record.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function

Function IsCursLocked(crsid)
 If UCase(Left(crsid,1)) = "L" then IsCursLocked = true else IsCursLocked = false
End Function

Function KillLockInfo(crsid)
 If IsCursLocked(crsid) then 
   KillLockInfo = Right(crsid, Len(crsid)-Len("L"))
 Else
   KillLockInfo = crsid
 End If  
End Function

' Evenimentul apare la apasarea butonului Register user at cours
Sub Button1_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  If IsCursLocked(RecList) Then
    MsgBox "You cannot un-enroll from a course were you are locked." & vbCrLf & "Please contact course professor.", vbOkOnly+vbExclamation
  Else
    if msgbox("Are you sure you want to un-enroll from selected course?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub
    FormularH.SelectAction.value = "unsubscr"
    FormularH.SelectList.value = RecList
    ActivateButtons false
    FormularH.submit
  End If  
End Sub


' Evenimentul apare la apasarea butonului Info curs
Sub Button2_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  ShowModalDialog "studcursinfo.asp?CursID=" & KillLockInfo(RecList),, "dialogWidth=450px;dialogHeight=311px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub  


' Evenimentul apare la apasarea butonului Enter cours
Sub Button3_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  If IsCursLocked(RecList) Then
    MsgBox "You cannot access selected course" & vbCrLf & "because the professor temporary blocked you.", vbOkOnly+vbExclamation
  Else
    HideAllDivs
    window.parent.frames("Header").location.href = "headerstud.asp?CursID=" & RecList
  End If  
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button4_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
