<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/settings.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
 
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("DSN")
	langOptions = GetOptionsForSelect("<option value='@0' @SELECTED>@1</option>", "@0 = " & GetPreferredLanguage(cn), "TBLanguages", Array("id", "langname"), cn)
	cn.Close
	Set cn = Nothing
%>
<html>
<head>
  <title>Add language</title>
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
<span id="Label1"
      class=TLabel style="width:70px;height:13px;"
      style="left:8px;top:24px;">
Language:
</span>
<input id="Edit1" type=text maxlength=50"
       class=TEdit style="width:185px;height:21px;"
       style="left:128px;top:16px;"
       title="Numele noii limbi">
<input CHECKED id=RadioButton1 type=radio name=langderiv
       class=TButton
       style="left:8px;top:64px;">
<label for=RadioButton1
       class=TLabel style="width:100px;height:13px;"
       style="cursor:hand; left:30px;top:68px;">Copy text from language
</label>
<input id=RadioButton2 type=radio name=langderiv
       class=TButton
       style="left:8px;top:136px;">
<label for=RadioButton2
       class=TLabel style="width:100px;height:13px;"
       style="cursor:hand; left:30px;top:140px;">Add empty texts
</label>
<select id="cboLanguageBase"
		class=TComboBox
		style="left:128px;top:88px;"
		style="width:185px;height:21px;">
		<%=langOptions%>
</select>
<input id="Button1" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:82px;top:176px;">
<input id="Button2" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:170px;top:176px;">
</div>

<script language=vbscript src="../_clientscripts/MiscControlUtils.vbs"></script>
<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
	Form1.style.visibility = "visible"
	WaitforForm.style.visibility = "hidden"
	Edit1.focus()
End Sub

' Evenimentul apare la apasarea butonului Cancel
Sub Button1_onclick
	Window.close 
End Sub

Sub RadioButton1_onclick
  TwoRadioOneEditDisable RadioButton2, cboLanguageBase
End Sub

Sub RadioButton2_onclick
  TwoRadioOneEditDisable RadioButton2, cboLanguageBase
End Sub

'Evenimentul apare la apasarea butonului OK
Sub Button2_onclick()
	Dim langname
	Dim baselang
	
	langname = Trim(Edit1.Value)
	If Len(langname) = 0 then
		msgbox "Language name cannot be empty.", vbOkOnly+vbExclamation
		Exit Sub
	ElseIf RadioButton1.checked and (cboLanguageBase.selectedIndex < 0) then
		msgbox "Language from were texts will be copied was not specified.", vbOkOnly+vbExclamation
		Exit Sub
	End If
    
    If RadioButton1.checked Then
		baselang = cboLanguageBase.value
	Else
		baselang = 0
	End If
    
	Window.Returnvalue = Array(langname, baselang)
	Window.Close
End Sub
</script>

</body>
</html>