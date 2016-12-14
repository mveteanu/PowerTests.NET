<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim ProfsList


Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
ProfsList = GetFillSelectFromDict(GetProfsDict(cn),"")
cn.Close
set cn = nothing

Function GetProfsDict(objCon)
 Dim re
 Set re = CreateObject("Scripting.Dictionary")
 Set rs = objCon.Execute("GetProfsInfoShort")
 do until rs.EOF
   re.Add CStr(rs.Fields("id_prof").value), rs.Fields("fullname").value
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set GetProfsDict = re
End Function
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:750px;height:308px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;" unselectable="on"
     class=TForm style="width:750px;height:308px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<div class=TForm style="width:725px;height:41px;" unselectable="on"
     style="left:8px;top:8px;">
<select id="ComboBox1"
        class=TComboBox style="width:700px;left:8px;top:10px;">
<option SELECTED value="-1">Select professor</option>
<option value="">All professors</option>
<%=ProfsList%>
</select>
</div>
<%CreateTableControl 8, 48, 217, Array("Course name", "Enrolled students", "Enrollment requests", "Maximum students", "Enrollment policy"), Array(195,130,130,130,140), 1, "studregcours_dat.asp" , true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Info course" title="Course information"
       class=TButton style="width:75px;height:25px;"
       style="left:246px;top:272px;">
<input DISABLED id="Button2" type=button value="Enroll" title="Enroll to selected course"
       class=TButton style="width:75px;height:25px;"
       style="left:334px;top:272px;">
<input id="Button3" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:422px;top:272px;">
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
Dim TDCFirtTime

' Schimba starea activ/inactiv a butoanelor
Sub ActivateButtons(btnsstate)
 Dim i
 
 for i=1 to 2
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
  If IsEmpty(TDCFirtTime) then
    TDCFirtTime = false
    tdcMyTable.filter = "id_prof=-1" 
    tdcMyTable.reset
  End If  
    ActivateButtons true
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub

' Apare la selectarea din ComboBox
Sub ComboBox1_OnChange
  tdcMyTable.filter = "id_prof=" & ComboBox1.value 
  tdcMyTable.reset
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




' Evenimentul apare la apasarea butonului Info curs
Sub Button1_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  ShowModalDialog "studcursinfo.asp?CursID=" & RecList,, "dialogWidth=450px;dialogHeight=311px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub  


' Evenimentul apare la apasarea butonului Register user at cours
Sub Button2_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  FormularH.SelectAction.value = "reg"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button3_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
