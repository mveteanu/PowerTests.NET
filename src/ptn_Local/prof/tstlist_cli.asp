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
     class=TForm style="width:770px;height:310px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;" unselectable='on'
     class=TForm style="width:770px;height:310px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 8, 8, 241, Array("Test name", "Time", "Use quest", "Not use quest", "Random q.", "Max Solvings", "Visible", "Solvings"), Array(135,60,90,110,90,105,65,95), 2, "tstlist_dat.asp" , true, "MyTable"%>
<button id="Button1" title="Selectare teste"
       class=TButton style="width:80px;height:25px;"
       style="left:8px;top:264px;">Select <font face="Webdings">6</font></button>
<input id="Button2" type=button value="Rename" title="Rename test"
       class=TButton style="width:80px;height:25px;"
       style="left:103px;top:264px;">
<input id="Button3" type=button value="Add" title="Add a new test"
       class=TButton style="width:80px;height:25px;"
       style="left:198px;top:264px;">
<input id="Button4" type=button value="Delete" title="Delete selected tests"
       class=TButton style="width:80px;height:25px;"
       style="left:293px;top:264px;">
<input id="Button5" type=button value="Edit" title="Edit selected test"
       class=TButton style="width:80px;height:25px;"
       style="left:387px;top:264px;">
<button id="Button6" title="Generate a test for viewing purposes"
       class=TButton style="width:80px;height:25px;"
       style="left:483px;top:264px;">View test <font face="Webdings">6</font></button>
<button id="Button7" title="Change test visibility"
       class=TButton style="width:80px;height:25px;"
       style="left:579px;top:264px;">Visibility <font face="Webdings">6</font></button>
<input id="Button8" type=button value="Close" title="Close form"
       class=TButton style="width:80px;height:25px;"
       style="left:675px;top:264px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="tstlist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
Dim mymenuview

MenuItems1 = Array("Select all","Unselect all")
MenuItems2 = Array("Screen preview","Print preview")
MenuItems3 = Array("Show to students","Hide from students")

' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub

' Activeaza/dezactiveaza butoanele
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


' Evenimentul apare la apasarea butonului Selectare
Sub Button1_OnClick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button1.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button1.style.top)  + StyleSizeToInt(Button1.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 140, "handlemenuviewclick", MenuItems1)
End Sub

' Evenimentul apare la apasarea butonului Rename
Sub Button2_OnClick
  Dim RecList, TstNewName
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  TstNewName = ShowModalDialog("tstren.asp?TstId=" & RecList, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If TstNewName<>"" then
   FormularH.SelectAction.value = "ren"
   FormularH.SelectList.value = RecList
   FormularH.SelectValue.value = TstNewName
   ActivateButtons false
   FormularH.submit 
  End If 
End Sub

' Evenimentul apare la apasarea butonului Add
Sub Button3_OnClick
  r = ShowModalDialog("tstcompose.asp", , "dialogWidth=649px;dialogHeight=367px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If r <> "" then 
   FormularH.SelectAction.value = "add"
   FormularH.SelectValue.value  = r
   ActivateButtons false
   FormularH.submit 
  End If
End Sub

' Evenimentul apare la apasarea butonului Delete
Sub Button4_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to delete selected tests?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub

  FormularH.SelectAction.value = "del"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub

' Handlerul executat la selectarea unei optiuni din meniul Select
Sub handlemenuviewclick(html)
 if html="<HR>" or html="" then exit sub
 mymenuview.Hide
 set mymenuview=nothing

 select case html     
  case MenuItems1(0) TableSelectAll tblMyTable, true
  case MenuItems1(1) TableSelectAll tblMyTable, false
  case MenuItems2(0) PreviewSampleTest false
  case MenuItems2(1) PreviewSampleTest true
  case MenuItems3(0) ChangeTestVisibility true
  case MenuItems3(1) ChangeTestVisibility false
 end select
End Sub

Sub PreviewSampleTest(laImprimanta)
  Dim RecList, ExtraStr, WinHeight
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  If laImprimanta then 
    ExtraStr  = "prn=yes&"
    WinHeight = 520
  else 
    ExtraStr = ""
    WinHeight = 500
  End If  
  ShowModalDialog "tstviewscrprn.asp?"& ExtraStr &"TstID=" & RecList, "", "dialogWidth=720px;dialogHeight="& WinHeight &"px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub

' Trimite catre server cererea de schimbare a vizibilitatii testului
Sub ChangeTestVisibility(vis)
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "pub"
  FormularH.SelectList.value = RecList
  FormularH.SelectValue.value = CStr(vis)
  ActivateButtons false
  FormularH.submit 
End Sub

' Evenimentul apare la apasarea butonului Edit test
Sub Button5_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  r = ShowModalDialog("tstcompose.asp?TstId=" & RecList, , "dialogWidth=649px;dialogHeight=367px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If r <> "" then 
   FormularH.SelectAction.value = "edt"
   FormularH.SelectList.value = RecList
   FormularH.SelectValue.value = r
   ActivateButtons false
   FormularH.submit 
  End If
End Sub

' Evenimentul apare la apasarea butonului View Generated Test
Sub Button6_onclick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button6.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button6.style.top)  + StyleSizeToInt(Button6.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 120, "handlemenuviewclick", MenuItems2)
End Sub


' Evenimentul apare la apasarea butonului Visibility
Sub Button7_onclick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button7.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button7.style.top)  + StyleSizeToInt(Button7.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 140, "handlemenuviewclick", MenuItems3)
End Sub

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button8_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
