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
     class=TForm style="width:545px;height:272px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:545px;height:272px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 7, 8, 201, Array("Category name",  "Questions"), Array(400,122), 2, "pbcateglist_dat.asp" , true, "MyTable"%>
<button DISABLED id="Button1" title="Selectare automata inregistrari"
       class=TButton style="width:75px;height:25px;"
       style="left:8px;top:224px;">Select <font face="Webdings">6</font></button>
<input DISABLED id="Button2" type=button value="Add" title="Add category"
       class=TButton style="width:75px;height:25px;"
       style="left:98px;top:224px;">
<input DISABLED id="Button3" type=button value="Delete" title="Delete category"
       class=TButton style="width:75px;height:25px;"
       style="left:187px;top:224px;">
<input DISABLED id="Button4" type=button value="Rename" title="Rename category"
       class=TButton style="width:75px;height:25px;"
       style="left:277px;top:224px;">
<input DISABLED id="Button5" type=button value="Edit" title="Edit category"
       class=TButton style="width:75px;height:25px;"
       style="left:366px;top:224px;">
<input DISABLED id="Button6" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:456px;top:224px;">
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="pbcateglist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>


<script language=vbscript>
Dim mymenuselect

MenuItems1 = Array("Select all","Unselect all")

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
    MsgBox "You need first to select a record.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function


' Handlerul meniului care apare la apasarea butonului Select
Sub handlemymenuselectwclick(html)
 Dim RecList

 if html="<HR>" or html="" then exit sub
 mymenuselect.Hide
 set mymenuselect=nothing

 select case html     
  case MenuItems1(0) TableSelectAll tblMyTable, true
  case MenuItems1(1) TableSelectAll tblMyTable, false
 end select
End Sub


' Evenimentul apare la apasarea butonului Select
Sub Button1_OnClick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button1.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button1.style.top)  + StyleSizeToInt(Button1.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuselect = showmenu(leftm, topm, 140, "handlemymenuselectwclick", MenuItems1)
End Sub


' Evenimentul apare la apasarea butonului Add category
Sub Button2_OnClick
  Dim CategNewName

  CategNewName = ShowModalDialog("pbtstcategadd.asp", , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If CategNewName<>"" then
   FormularH.SelectAction.value = "add"
   FormularH.SelectValue.value = CategNewName
   ActivateButtons false
   FormularH.submit 
  End If 
End Sub

' Evenimentul apare la apasarea butonului Delete category
Sub Button3_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to delete selected categories?",vbYesNo+vbQuestion,"Confirm") = vbNo then
    Exit Sub
  End If     

  FormularH.SelectAction.value = "del"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Rename
Sub Button4_OnClick
  Dim RecList, CategNewName
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  CategNewName = ShowModalDialog("pbtstcategadd.asp?CategID=" & RecList, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If CategNewName<>"" then
   FormularH.SelectAction.value = "ren"
   FormularH.SelectList.value = RecList
   FormularH.SelectValue.value = CategNewName
   ActivateButtons false
   FormularH.submit 
  End If 
End Sub


' Evenimentul care apare la apasarea butonului Edit pb. in categ
Sub Button5_OnClick
  Dim RecList, PBIDs
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  PBIDs = ShowModalDialog("pbcategedt.asp?CategID=" & RecList, , "dialogWidth=506px;dialogHeight=375px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If PBIDs = "" then Exit Sub
  
  If PBIDs = "<vid>" then PBIDs = ""
  FormularH.SelectAction.value = "edt"
  FormularH.SelectList.value = RecList
  FormularH.SelectValue.value = PBIDs
  ActivateButtons false
  FormularH.submit 
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button6_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
