<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/cursuri.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
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

If Request.QueryString.Count>0 then
  TitluFereastra = "Edit course"
  TitluOKBtn = "Save"
  set cn=Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  set rsc = GetCursByID(Request.QueryString("cursid"),cn)
  Curs_Nume = rsc.Fields("numecurs").value
  Curs_Public = rsc.Fields("curspublic").value
  Curs_Permisii = rsc.Fields("permisiiacceptare").value
  Curs_MaxStud = rsc.Fields("maxstudents").value
  rsc.Close
  set rsc=nothing
  cn.Close 
  set cn=nothing
else
  TitluFereastra = "Add course"
  TitluOKBtn = "Add course"
  Curs_Nume = ""
  Curs_Public = true
  Curs_Permisii = 0
  Curs_MaxStud = 0
end if
Curs_MaxStudNelim = CBool(Curs_MaxStud = 0)
%>
<html>
<head>
  <title><%=TitluFereastra%></title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

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
<fieldset class=TGroupBox style="width:321px;height:137px;"
          style="left:8px;top:8px;">
<legend>Course properties</legend>
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
</fieldset>
<fieldset class=TGroupBox style="width:321px;height:153px;"
          style="left:8px;top:152px;">
<legend>Students enrollment policy</legend>
<span id="Label3"
      class=TLabel style="width:89px;height:13px;"
      style="left:16px;top:40px;">
Enrollment policy:
</span>
<span id="Label4"
      class=TLabel style="width:104px;height:13px;"
      style="left:16px;top:80px;">
Number of students:
</span>
<select id="ComboBox1"
        class=TComboBox style="width:185px;"
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
       class=TEdit style="width:109px;height:21px;"
       style="left:196px;top:110px;">
</fieldset>
<input id="Button1" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:176px;top:328px;">
<input id="Button2" type=button value="<%=TitluOKBtn%>" title="<%=TitluOKBtn%>"
       class=TButton style="width:75px;height:25px;"
       style="left:88px;top:328px;">
</div>


<script language=vbscript src="../_clientscripts/MiscControlUtils.vbs"></script>
<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  Edit1.focus 
End Sub


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


Sub RadioButton3_onclick
  TwoRadioOneEditDisable RadioButton3, Edit2
End Sub

Sub RadioButton4_onclick
  TwoRadioOneEditDisable RadioButton3, Edit2
End Sub

' Evenimentul apare la apasarea butonului Cancel
Sub Button1_onclick
  Window.close 
End Sub


Sub Button2_onclick
  if Edit1.value = "" then 
    msgbox "You need to enter course name.",vbOkOnly+vbExclamation
  elseif (RadioButton4.checked = true) and (not IsNumeric(Edit2.value)) then
    msgbox "Maximum number of students should be a numeric value.",vbOkOnly+vbExclamation
  else
    Window.Returnvalue = GetCursData()
    Window.Close
  end if
End Sub

</script>

</body>
</html>