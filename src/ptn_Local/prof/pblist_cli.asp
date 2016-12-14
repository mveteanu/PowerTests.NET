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
     class=TForm style="width:690px;height:310px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;" unselectable='on'
     class=TForm style="width:690px;height:310px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 8, 8, 241, Array("Question name", "Answer type", "Partial answer", "Answers", "Images", "Categories", "Tests"), Array(120,95,110,95,75,95,75), 2, "pblist_dat.asp" , true, "MyTable"%>
<button id="Button1" title="Select/Unselect all"
       class=TButton style="width:85px;height:25px;"
       style="left:8px;top:264px;">Select <font face="Webdings">6</font></button>
<input id="Button2" type=button value="Rename" title="Rename question"
       class=TButton style="width:85px;height:25px;"
       style="left:105px;top:264px;">
<input id="Button3" type=button value="Add" title="Add a new question"
       class=TButton style="width:85px;height:25px;"
       style="left:201px;top:264px;">
<input id="Button4" type=button value="Delete" title="Delete selected questions"
       class=TButton style="width:85px;height:25px;"
       style="left:298px;top:264px;">
<button id="Button5" title="View selected questions"
       class=TButton style="width:85px;height:25px;"
       style="left:395px;top:264px;">View <font face="Webdings">6</font></button>
<input id="Button6" type=button value="Edit" title="Edit selected question"
       class=TButton style="width:85px;height:25px;"
       style="left:491px;top:264px;">
<input id="Button7" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:588px;top:264px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="pblist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
Dim mymenuview

MenuItems1 = Array("Screen preview","Print preview")
MenuItems2 = Array("Select all","Unselect all")

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
 set mymenuview = showmenu(leftm, topm, 140, "handlemenuviewclick", MenuItems2)
End Sub

' Evenimentul apare la apasarea butonului Rename problem
Sub Button2_OnClick
  Dim RecList, PBNewName
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  PBNewName = ShowModalDialog("pbproprpb.asp?PBId=" & RecList, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If PBNewName<>"" then
   FormularH.SelectAction.value = "ren"
   FormularH.SelectList.value = RecList
   FormularH.SelectValue.value = PBNewName
   ActivateButtons false
   FormularH.submit 
  End If 
End Sub

' Evenimentul apare la apasarea butonului Add problem
Sub Button3_OnClick
  r = CInt(ShowModalDialog("pbcomposepb.asp", , "dialogWidth=768px;dialogHeight=547px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"))
  If r = 1 then ReloadTDC
End Sub

' Evenimentul apare la apasarea butonului Delete problem
Sub Button4_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to delete selected questions?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub

  FormularH.SelectAction.value = "del"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub

' Evenimentul apare la apasarea butonului View problem
Sub Button5_onclick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button5.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button5.style.top)  + StyleSizeToInt(Button5.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuview = showmenu(leftm, topm, 120, "handlemenuviewclick", MenuItems1)
End Sub

' Handlerul executat la selectarea unei optiuni din meniul View
Sub handlemenuviewclick(html)
 Dim RecList

 if html="<HR>" or html="" then exit sub
 mymenuview.Hide
 set mymenuview=nothing

 select case html     
  case MenuItems2(0) TableSelectAll tblMyTable, true
  case MenuItems2(1) TableSelectAll tblMyTable, false
  case MenuItems1(0) RecList=GetSelectedRecords
                     if RecList="" then Exit Sub
                     ShowModalDialog "pbviewscr.asp?PBIDs=" & RecList, "", "dialogWidth=720px;dialogHeight=500px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
  case MenuItems1(1) RecList=GetSelectedRecords
                     if RecList="" then Exit Sub 
                     ShowModalDialog "pbviewprn.asp?PBIDs=" & RecList,, "dialogWidth=720px;dialogHeight=520px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
 end select
End Sub


' Evenimentul apare la apasarea butonului Edit problem
Sub Button6_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  r = CInt(ShowModalDialog("pbcomposepb.asp?PBId=" & RecList, , "dialogWidth=768px;dialogHeight=547px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"))
  If r = 1 then ReloadTDC
End Sub


' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button7_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
