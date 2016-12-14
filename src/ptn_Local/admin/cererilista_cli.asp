<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 tipuser=Request.QueryString("tipuser")
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:670px;height:350px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:670px;height:350px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 12, 8, 285, Array("Last name", "First name", "Email", "Phone", "Login", "Sign-up date"), Array(100,100,145,100,80,120), 2, "cererilista_dat.asp?tipuser="& CStr(tipuser) , true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Select all" title="Select all records"
       class=TButton style="width:110px;height:25px;"
       style="left:12px;top:308px;">
<input DISABLED id="Button2" type=button value="Unselect all" title="Unselect all records"
       class=TButton style="width:110px;height:25px;"
       style="left:146px;top:308px;">
<input DISABLED id="Button3" type=button value="Accept selected" title="Grant access to selected persons"
       class=TButton style="width:110px;height:25px;"
       style="left:280px;top:308px;">
<input DISABLED id="Button4" type=button value="Deny selected" title="Deny access to selected persons"
       class=TButton style="width:110px;height:25px;"
       style="left:413px;top:308px;">
<input id="Button5" type=button value="Close" title="Close form"
       class=TButton style="width:110px;height:25px;"
       style="left:547px;top:308px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="cererilista_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
' Schimba starea activ/inactiv a celor butoanelor
Sub ActivateButtons(b1, b2, b3, b4, b5)
  Button1.disabled = not b1
  Button2.disabled = not b2
  Button3.disabled = not b3
  Button4.disabled = not b4
  Button5.disabled = not b5
End Sub

' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub


' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  ActivateButtons true, true, true, true, true
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
  if RecList = "" then  MsgBox "No record selected.", vbOkOnly+vbExclamation
  GetSelectedRecords = RecList
End Function


' Evenimentul apare la apasarea butonului Select All
Sub Button1_OnClick
  TableSelectAll tblMyTable, true
End Sub

' Evenimentul apare la apasarea butonului Select none
Sub Button2_OnClick
  TableSelectAll tblMyTable, false
End Sub


' Evenimentul apare la apasarea butonului Valideaza selectati
Sub Button3_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub

  FormularH.SelectAction.value = "accept"
  FormularH.SelectList.value = RecList
  ActivateButtons false, false, false, false, false
  FormularH.submit 
End Sub


' Evenimentul apare la apasarea butonului Respinge selectati
Sub Button4_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Are you sure you want to deny access to selected users?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub

  FormularH.SelectAction.value = "deny"
  FormularH.SelectList.value = RecList
  ActivateButtons false, false, false, false, false
  FormularH.submit 
End Sub


' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button5_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
