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

<script language="vbscript" src="../_clientscripts/utils.vbs"></script>
<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:691px;height:330px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:691px;height:330px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 8, 8, 273, Array("Last name", "First name", "Email", "Phone", "Login", "No of testings", "Locked"), Array(105,100,100,90,80,120,70), 2, "studlist_dat.asp" , true, "MyTable"%>
<button DISABLED id="Button1" title="Select records"
       class=TButton style="width:85px;height:25px;"
       style="left:8px;top:288px;">Select <font face="Webdings">6</font></button>
<input DISABLED id="Button2" type=button value="Lock" title="Lock selected students access to this course"
       class=TButton style="width:85px;height:25px;"
       style="left:105px;top:288px;">
<input DISABLED id="Button3" type=button value="Unlock" title="Unlock selected students"
       class=TButton style="width:85px;height:25px;"
       style="left:203px;top:288px;">
<input DISABLED id="Button4" type=button value="Un-enroll" title="Un-enroll selected students"
       class=TButton style="width:85px;height:25px;"
       style="left:300px;top:288px;">
<input DISABLED id="Button5" type=button value="Info" title="Student summary"
       class=TButton style="width:85px;height:25px;"
       style="left:397px;top:288px;">
<button DISABLED id="Button6" title="See student test results"
       class=TButton style="width:85px;height:25px;"
       style="left:495px;top:288px;">Results <font face="Webdings">6</font></button>
<input id="Button7" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:592px;top:288px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="studlist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
Dim mymenuview

MenuItems1 = Array("Select all","Unselect all")
MenuItems2 = Array("Results for all tests","Results for a test","<HR>","Results of a student")

' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub

' Schimba starea activ/inactiv a celor butoanelor
Sub ActivateButtons(state)
 For i = 1 to 7
   Form1.all("Button" & CStr(i)).disabled = not state
 Next
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

' Intoarce in format CSV ID-urile recordurilor selectate si afiseaza un
' mesaj de avertizare daca nu s-a selectat nici o inregistrare
Function GetSelectedRecords
  Dim RecList
  RecList = TableGetSelected(tblMyTable)
  if RecList = "" then  MsgBox "You need to select at least one record.", vbOkOnly+vbExclamation
  GetSelectedRecords = RecList
End Function


' Intoarce sub forma de string ID-ul recordului selectat.
' Daca nu se selecteaza nici o inregistrare sau se selecteaza mai mult de una
' atunci se afiseaza un mesaj si se intoarce sirul vid.
Function GetSelectedRecord
  Dim RecList
  Dim RecArray
  
  RecList  = TableGetSelected(tblMyTable)
  RecArray = Split(RecList,",",-1,1)
  If (UBound(RecArray)-LBound(RecArray))<>0 then 
    MsgBox "You need to select a single record.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function

' Handlerul executat la selectarea unei optiuni din meniul View
Sub handlemenuviewclick(html)
 Dim RecList

 if html="<HR>" or html="" then exit sub
 mymenuview.Hide
 set mymenuview=nothing

 select case html     
  case MenuItems1(0) TableSelectAll tblMyTable, true
  case MenuItems1(1) TableSelectAll tblMyTable, false
  
  case MenuItems2(0) ShowModalDialog "studclasament1.asp",, "dialogWidth=620px;dialogHeight=520px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
  case MenuItems2(1) ShowModalDialog "seltstrez_cli.asp",, "dialogWidth=475px;dialogHeight=345px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
  case MenuItems2(3) RecList=GetSelectedRecord
                     if RecList="" then Exit Sub    
                     ShowModalDialog "../stud/studseltstrez_cli.asp?userid=" & GetTDCData(tdcMyTable, RecList, Array("IDStud"))(0),, "dialogWidth=500px;dialogHeight=387px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
 end select
End Sub

' Evenimentul apare la apasarea butonului Select
Sub Button1_onclick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button1.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button1.style.top)  + StyleSizeToInt(Button1.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 140, "handlemenuviewclick", MenuItems1)
End Sub

' Evenimentul apare la apasarea butonului Rezultate
Sub Button6_onclick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button6.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button6.style.top)  + StyleSizeToInt(Button6.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 180, "handlemenuviewclick", MenuItems2)
End Sub


' Evenimentul care apare la apasarea butonului Lock
Sub Button2_onclick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "lock"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului UnLock
Sub Button3_onclick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "unlock"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Da afara
Sub Button4_onclick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to un-enroll selected students?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub

  FormularH.SelectAction.value = "delsubscript"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub

' Evenimentul apare la apasarea butonului Info student
Sub Button5_onclick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  ShowModalDialog "studinfo.asp?StudID=" & RecList, , "dialogWidth=497px;dialogHeight=231px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button7_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
