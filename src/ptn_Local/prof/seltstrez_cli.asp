<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
%>
<HTML>
<head>
 <title>Select a test</title>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language="vbscript" src="../_clientscripts/utils.vbs"></script>
<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<%
 'Nume test|Nr rezolvari|Nr abandonari
 CreateTableControl 8, 8, 259, Array("Test name","Completed","Abandoned"), Array(229,110,110), 1, "seltstrez_dat.asp", true, "MyTable" 
%>

<input DISABLED id="Button1" type=button value="View results" title="View results for this test"
       class=TButton style="width:120px;height:25px;"
       style="left:112px;top:280px;">
<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:120px;height:25px;"
       style="left:260px;top:280px;">
</div>

<script language=vbscript>
' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub

' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  Button1.disabled = false
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  Window.close   
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
    MsgBox "You need to select a single record.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function

' Evenimentul apare la apasarea butonului Show results
Sub Button1_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  ShowModalDialog "studclasament1.asp?TestID=" & RecList,, "dialogWidth=620px;dialogHeight=520px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button2_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
