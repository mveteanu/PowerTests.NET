<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/settings.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<%
Function LanguageOptions
	Dim re
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("DSN")
	re = GetOptionsForSelect("<option value='@0' @SELECTED>@1</option>", "@0 = " & GetPreferredLanguage(cn), "TBLanguages", Array("id", "langname"), cn)
	cn.Close
	Set cn = Nothing
	LanguageOptions = re
End Function
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<form ID="RegForm" action="reguser_ifr.asp" method="post" target=ConfirmIframe>
<div id="Form1" unselectable="on"
     class=TForm style="width:350px;height:348px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:60px;">
<span unselectable="on" class=TLabel style="width:119px;height:13px;font-weight:bold;"
      style="left:117px;top:10px;">
New user sign-up
</span>

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:56px;">
Account type *:
</span>
<select name=tipuser id=tipuser
		class=TComboBox
		style="left:126px;top:48px;"
		style="width:180px;height:21px;">
		<option value="A">Administrator</option>
		<option value="P">Professor</option>
		<option value="S" selected>Student</option>
</select>

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:84px;">
Login *:
</span>
<input autocomplete="off" name=login id=login
       type=text maxlength=20 class=TEdit style="left:126px;top:76px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:112px;">
Password *:
</span>
<input autocomplete="off" name=pass id=pass
       type=password maxlength=20 class=TEditPassword style="left:126px;top:104px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:140px;">
Verify password *:
</span>
<input autocomplete="off" name=passagain id=passagain
       type=password maxlength=20 class=TEditPassword style="left:126px;top:132px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:168px;">
Last name *:
</span>
<input autocomplete="off" name=nume id=nume
       type=text maxlength=20 class=TEdit style="left:126px;top:160px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:196px;">
First name *:
</span>
<input autocomplete="off" name=prenume id=prenume
       type=text maxlength=20 class=TEdit style="left:126px;top:188px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:224px;">
Email *:
</span>
<input autocomplete="off" name=email id=email
       type=text maxlength=20 class=TEdit style="left:126px;top:216px;"
       style="width:180px;height:21px;">

<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:252px;">
Phone:
</span>
<input autocomplete="off" name=telefon id=telefon
       type=text maxlength=20 class=TEdit style="left:126px;top:244px;"
       style="width:180px;height:21px;">


<span unselectable="on" class=TLabel style="width:100px;height:13px;"
      style="left:20px;top:280px;">
Preferred language:
</span>
<select id="cboLanguage" name="cboLanguage"
		class=TComboBox
		style="left:126px;top:272px;"
		style="width:180px;height:21px;">
		<%=LanguageOptions()%>
</select>

<input id="BtnOK" type=submit value="Sign-up" title="Sign-up"
	   class=TButton style="width:75px;height:25px;"
       style="left:110px;top:308px;">
<input id="BtnCancel" type=button value="Cancel" title="Cancel sign-up"
       class=TButton style="width:75px;height:25px;"
       style="left:190px;top:308px;">

<div id="Form2" style="display: none; position:absolute; border:groove thin;"
                style="left:2px; top:308px; width:342px; height:120px; overflow:hidden;">
<IFRAME ID=ConfirmIframe Name=ConfirmIframe FRAMEBORDER=No FRAMESPACING=0 width=100% height=100% scrolling=no>
</IFRAME>
</div>

</div>
</form>

<script language=vbscript>
' Event handler for click on Cancel button
Sub btnCancel_OnClick
	window.location = "login.asp"
End Sub
</script>


<script language=JavaScript>
// Validates email addresses
function validEMail(a)
{
	var pos1 = a.lastIndexOf('@');
	var pos2 = a.indexOf('.');

	return ( (pos1 != -1) && (pos2 != -1) && (pos1 != 0) && (pos1 != a.length) && 
		(pos2 != 0) && (pos2 != a.length) && (pos1 == a.indexOf('@')) && 
		(a.charAt(pos1-1) != ' ') && (a.charAt(pos1+1) != ' ') && 
		(a.charAt(pos2-1) != ' ') && (a.charAt(pos2+1) != ' ') );
}
</script>

<script language=vbscript>
' Check the user data before sending to the server
Function RegForm_OnSubmit
	If Len(RegForm.login.value) = 0 Then
		MsgBox "Login cannot be empty", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.login.focus
	ElseIf Len(RegForm.pass.value) = 0 Then
		MsgBox "Password cannot be empty.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.pass.focus
	ElseIf Len(RegForm.passagain.value) = 0 Then
		MsgBox "Password verification cannot be empty.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.passagain.focus
	ElseIf RegForm.pass.value <> RegForm.passagain.value Then
		MsgBox "Password verification is different than password.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.passagain.focus
	ElseIf Len(RegForm.nume.value) = 0 Then
		MsgBox "Last name cannot be empty. Please use your real name as data will be validated.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.nume.focus
	ElseIf Len(RegForm.prenume.value) = 0 Then
		MsgBox "First name cannot be empty. Please use your real name as data will be validated.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.prenume.focus
	ElseIf Not validEMail(RegForm.email.value) Then
		MsgBox "Email address is invalid. Please use a valid email address.", vbExclamation, "PowerTests .NET"
		RegForm_onSubmit = false
		RegForm.email.focus
	ElseIf RegForm.cboLanguage.selectedIndex = -1 Then
		MsgBox "Please select your preferred language."
		RegForm_onSubmit = false
	End If 
End Function


' Displays the DIV that contains the IFRAME where 
' the response of username and password form will be loaded
Public Sub ShowConfirm()
	RegForm.BtnOK.style.visibility = "hidden"
	RegForm.BtnCancel.style.visibility = "hidden"
	Form1.style.height = "434px"
	Form2.style.display = "block"
	document.frames(0).focus 
End Sub


' Hides the DIV...
Public Sub HideConfirm()
	Form1.style.height = "348px"
	Form2.style.display = "none"
	With RegForm
		.BtnOK.style.visibility = "visible"
		.BtnOK.disabled = False
		.BtnCancel.style.visibility = "visible"
		.BtnCancel.disabled = False
	End With
End Sub


Public Sub ClearForm()
	With RegForm
		.tipuser.value = "S"
		.login.value = ""
		.pass.value = ""
		.passagain.value = ""
		.nume.value = ""
		.prenume.value = ""
		.email.value = ""
		.telefon.value = ""
	End With
End Sub
</script>


</BODY>
</HTML>
