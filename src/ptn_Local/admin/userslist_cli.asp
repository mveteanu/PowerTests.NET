<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
 Dim infodialogurl, infodialogw, infodialogh, tipuser
 
 Response.Buffer = True
 Response.Expires = -1

 tipuser=Request.QueryString("tipuser")
 Select case UCase(tipuser)
   case "A" infodialogurl = "usersinfoa_cli.asp"
            infodialogw = 318
            infodialogh = 282
   case "P" infodialogurl = "usersinfop_cli.asp"
            infodialogw = 593
            infodialogh = 308
   case "S" infodialogurl = "usersinfos_cli.asp"
            infodialogw = 593
            infodialogh = 282
 End Select
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
     class=TForm style="width:780px;height:348px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:780px;height:348px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%
select case UCase(tipuser)
 case "A" CreateTableControl 12, 8, 285, Array("Last name", "First name", "Login", "Email", "Phone", "Account Date","Locked"), Array(120,120,100,154,100,76,80), 2, "userslist_dat.asp?tipuser=A" , true, "MyTable"
 case "P" CreateTableControl 12, 8, 285, Array("Last name", "First name", "Login", "Email", "Phone", "Courses", "Students", "Account Date", "Locked"), Array(93,93,80,134,80,70,70,70,60), 2, "userslist_dat.asp?tipuser=P" , true, "MyTable"
 case "S" CreateTableControl 12, 8, 285, Array("Last name", "First name", "Login", "Email", "Phone", "Courses", "Account Date", "Locked"), Array(98,98,80,134,80,90,100,70), 2, "userslist_dat.asp?tipuser=S" , true, "MyTable"
end select
%>
<button DISABLED id="Button1" title="Selectare automata inregistrari"
       class=TButton style="width:75px;height:25px;"
       style="left:12px;top:304px;">Select <font face="Webdings">6</font></button>
<input DISABLED id="Button3" type=button value="Lock" title="Lock selected"
       class=TButton style="width:75px;height:25px;"
       style="left:96px;top:304px;">
<input DISABLED id="Button4" type=button value="Unlock" title="Unlock selected accounts"
       class=TButton style="width:75px;height:25px;"
       style="left:181px;top:304px;">
<input DISABLED id="Button5" type=button value="Delete" title="Delete selected accounts"
       class=TButton style="width:75px;height:25px;"
       style="left:265px;top:304px;">
<input DISABLED id="Button6" type=button value="Edit" title="Edit selected user"
       class=TButton style="width:75px;height:25px;"
       style="left:349px;top:304px;">
<input DISABLED id="Button7" type=button value="Info" title="Selected user info"
       class=TButton style="width:75px;height:25px;"
       style="left:433px;top:304px;">
<input DISABLED id="Button8" type=button value="Send email" title="Send email to selected users"
       class=TButton style="width:75px;height:25px;"
       style="left:518px;top:304px;">
<input DISABLED id="Button9" type=button value="Export" title="Export to Excel"
       class=TButton style="width:75px;height:25px;"
       style="left:602px;top:304px;">
<input id="Button10" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:686px;top:304px;">
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="userslist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValues" name="SelectValues">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>


<script language=vbscript>
Dim mymenuselect
MenuItems1 = Array("Select all","Unselect all")

' Schimba starea activ/inactiv a butoanelor
Sub ActivateButtons(b1, b3, b4, b5, b6, b7, b8, b9, b10)
  Button1.disabled = not b1
  Button3.disabled = not b3
  Button4.disabled = not b4
  Button5.disabled = not b5
  Button6.disabled = not b6
  Button7.disabled = not b7
  Button8.disabled = not b8
  Button9.disabled = not b9
  Button10.disabled = not b10
End Sub


' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub


' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  ActivateButtons true, true, true, true, true, true, true, true, true
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


' Handlerul executat la selectarea unei optiuni din meniul Select
Sub handlemenuselect(html)
 if html="<HR>" or html="" then exit sub
 mymenuselect.Hide
 set mymenuselect = nothing

 select case html     
  case MenuItems1(0) TableSelectAll tblMyTable, true
  case MenuItems1(1) TableSelectAll tblMyTable, false
 end select 
End Sub

' Evenimentul apare la apasarea butonului Select All
Sub Button1_OnClick
 Dim leftm, topm
   
 leftm = 2 + StyleSizeToInt(Button1.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 2 + StyleSizeToInt(Button1.style.top)  + StyleSizeToInt(Button1.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuselect = showmenu(leftm, topm, 140, "handlemenuselect", MenuItems1)
End Sub

' Evenimentul care apare la apasarea butonului Lock Selected
Sub Button3_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "lock"
  FormularH.SelectList.value = RecList
  ActivateButtons false, false, false, false, false, false, false, false, false
  FormularH.submit 
End Sub

' Evenimentul care apare la apasarea butonului UnLock Selected
Sub Button4_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "unlock"
  FormularH.SelectList.value = RecList
  ActivateButtons false, false, false, false, false, false, false, false, false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Delete Selected
Sub Button5_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to delete selected accounts?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub

  FormularH.SelectAction.value = "delete"
  FormularH.SelectList.value = RecList
  ActivateButtons false, false, false, false, false, false, false, false, false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Edit selected
Sub Button6_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  DatePers = ShowModalDialog("userslist_cliedt.asp?userid=" & RecList,, "dialogWidth=328px;dialogHeight=379px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  if not IsArray(DatePers) then Exit Sub

  FormularH.SelectAction.value = "edit"
  FormularH.SelectList.value = RecList
  FormularH.SelectValues.value = DatePers(0) & "|" & DatePers(1) & "|" & DatePers(2) & "|" & DatePers(3) & "|" & DatePers(4) & "|" & DatePers(5)
  ActivateButtons false, false, false, false, false, false, false, false, false
  FormularH.submit 
End Sub


' Evenimentul care apare la apasarea butonului Info cont
Sub Button7_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  ShowModalDialog "<%=infodialogurl%>?userid=" & RecList,, "dialogWidth=<%=infodialogw%>px;dialogHeight=<%=infodialogh%>px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub


' Evenimentul care apare la apasarea butonului Send email to selected
Sub Button8_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  ShowModalDialog "userseml_cli.asp?userlist=" & RecList,, "dialogWidth=500px;dialogHeight=450px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub


' Evenimentul care apare la apasarea butonului Export Selected
Sub Button9_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  'msgbox RecList
  msgbox "Excel export is not yet implemented"
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button10_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>

