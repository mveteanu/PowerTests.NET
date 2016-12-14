<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/cursuri.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #INCLUDE FILE="../_serverscripts/TabControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim TitluFereastra
Dim TitluOKBtn
Dim Curs_Nume
Dim Curs_Public
Dim Curs_Permisii
Dim Curs_MaxStud
Dim Curs_MaxStudNelim

set cn=Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
set rsc = GetCursByID(Session("CursID"),cn)
Curs_Nume = rsc.Fields("numecurs").value
Curs_Public = rsc.Fields("curspublic").value
Curs_Permisii = rsc.Fields("permisiiacceptare").value
Curs_MaxStud = rsc.Fields("maxstudents").value
rsc.Close
set rsc=nothing
cn.Close 
set cn=nothing

If Curs_MaxStud = 0 then 
  Curs_MaxStudNelim = true
Else
  Curs_MaxStudNelim = false
End If  
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
 <script language=javascript src="../_clientscripts/tabControlEvents.js"></script>
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:400px;height:264px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:400px;height:264px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%OpenTabControl 8,8,380, 200, Array("Properties", "Students"), 1, "PageControl1"%>
<%OpenTabContent%>
<span id="Label1"
      class=TLabel style="width:31px;height:13px;"
      style="left:16px;top:40px;">
Name:
</span>
<span id="Label2"
      class=TLabel style="width:32px;height:13px;"
      style="left:16px;top:80px;">
Public:
</span>
<input id="Edit1" type=text maxlength=50 value="<%=Curs_Nume%>"
       class=TEdit style="width:185px;height:21px;"
       style="left:120px;top:32px;">
<input <%=CompareToString(Curs_Public,true,"CHECKED","")%> id="RadioButton1" name=radiogrp1 type=radio value="0"
       class=TButton
       style="left:120px;top:80px;">
<input <%=CompareToString(Curs_Public,false,"CHECKED","")%> id="RadioButton2" name=radiogrp1 type=radio value="1"
       class=TButton
       style="left:120px;top:104px;">
<label for=RadioButton1 
       class=TLabel style="width:32px;height:13px;"
       style="cursor:hand; left:142px;top:84px;">Yes
</label>
<label for=RadioButton2 
       class=TLabel style="width:32px;height:13px;"
       style="cursor:hand; left:142px;top:108px;">No
</label>
<%CloseTabContent%>
<%OpenTabContent%>
<span id="Label3"
      class=TLabel style="width:89px;height:13px;"
      style="left:16px;top:40px;">
Enrollment policy:
</span>
<span id="Label4"
      class=TLabel style="width:74px;height:13px;"
      style="left:16px;top:80px;">
Students:
</span>
<select id="ComboBox1"
        class=TComboBox style="width:197px;"
        style="left:120px;top:36px;">
<option <%=CompareToString(Curs_Permisii,0,"SELECTED","")%> value=0>Automatic enrollment</option>
<option <%=CompareToString(Curs_Permisii,1,"SELECTED","")%> value=1>Needs validation</option>
<option <%=CompareToString(Curs_Permisii,2,"SELECTED","")%> value=2>Automatic deny</option>
</select>
<input <%=CompareToString(Curs_MaxStudNelim,true,"CHECKED","")%> id="RadioButton3" name=radiogrp2 type=radio value="0"
       class=TButton
       style="left:120px;top:80px;">
<input <%=CompareToString(Curs_MaxStudNelim,false,"CHECKED","")%> id="RadioButton4" name=radiogrp2 type=radio value="1"
       class=TButton
       style="left:120px;top:112px;">
<label for=RadioButton3 
       class=TLabel style="width:60px;height:13px;"
       style="cursor:hand; left:142px;top:84px;">Unlimited
</label>
<label for=RadioButton4 
       class=TLabel style="width:60px;height:13px;"
       style="cursor:hand; left:142px;top:116px;">Maximum
</label>
<input <%=CompareToString(Curs_MaxStudNelim,true,"DISABLED","")%> id="Edit2" type=text maxlength=20 value="<%=CompareToString(Curs_MaxStudNelim,true,"",CStr(Curs_MaxStud))%>"
       class=TEdit style="width:121px;height:21px;"
       style="left:196px;top:110px;">
<%CloseTabContent%>
<%CloseTabControl%>

<input id="Button1" type=button value="Save" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:120px;top:220px;">
<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:205px;top:220px;">
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
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  Edit1.focus 
End Sub


' Schimba starea activ/inactiv a butoanelor de Save si Close
' Daca state = true butoanele sunt active si invers
Sub ActivateButtons(state)
  Button1.disabled = not state
  Button2.disabled = not state
End Sub



' Atentie: Aceasta subrutina nu incarca nici un TDC! 
' Ea a trebuit definita pentru a se putea folosi aceasta pagina
' impreuna cu partea de server "cursurilist_ser.asp" !!!
Sub ReloadTDC
  ActivateButtons true
End Sub


' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close. 
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Trateaza radio-button-urile care comuta intre Nelimitat/Maxim
Sub RadioButton3_onclick
  if RadioButton3.checked then 
    Edit2.disabled = true
  else
    Edit2.disabled = false
    Edit2.focus
  end if    
End Sub

Sub RadioButton4_onclick
  RadioButton3_onclick
End Sub


' Obtine un array cu valorile care se afla in form
Function GetCursData
  Dim CursData(4)

  CursData(0) = Edit1.value        ' nume
  if RadioButton3.checked then
    CursData(1) = "0"              ' maxstudents
  else
    CursData(1) = Edit2.value
  end if
  CursData(2) = ComboBox1.Value    ' permisiiacceptare
  CursData(3) = CStr(CInt(RadioButton1.checked)) ' public
  
  GetCursData = CursData
End Function


' Evenimentul apare la apasarea butonului Close
Sub Button2_OnClick
  HideAllDivs
End Sub

' Evenimentul apare la apasarea butonului Save
Sub Button1_OnClick
  if Edit1.value = "" then 
    msgbox "You need to enter course name.",vbOkOnly+vbExclamation
  elseif (RadioButton4.checked = true) and (not IsNumeric(Edit2.value)) then
    msgbox "Maximum number of students should be a numeric value.",vbOkOnly+vbExclamation
  else
    DateCurs = GetCursData()
    
    FormularH.SelectAction.value = "edit"
    FormularH.SelectList.value = "<%=Session("CursID")%>"
    FormularH.SelectValues.value = DateCurs(0) & "|" & DateCurs(1) & "|" & DateCurs(2) & "|" & DateCurs(3)
    ActivateButtons false
    FormularH.submit 
  end if
End Sub
</script>

</body>
</html>
