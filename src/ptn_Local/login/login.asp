<%@ Language=VBScript %>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<form ID="LoginForm" action="login_ifr.asp" method="post" target=ConfirmIframe>
<div id="Form1" unselectable="on"
     class=TForm style="width:400px;height:180px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<span unselectable="on" class=TLabel style="width:119px;height:13px;font-weight:bold;"
      style="left:137px;top:4px;">
Welcome to PTN
</span>
<span unselectable="on" class=TLabel style="width:31px;height:13px;"
      style="left:88px;top:56px;">
Login:
</span>
<span unselectable="on" class=TLabel style="width:33px;height:13px;"
      style="left:86px;top:84px;">
Pass:
</span>
<input autocomplete="off" ID="LoginUser" name="LoginUser"
	   type=text class=TEdit style="left:126px;top:48px;"
       style="width:180px;height:21px;">
<input autocomplete="off" ID="LoginPass" name="LoginPass"
       type=password class=TEditPassword style="left:126px;top:76px;"
       style="width:180px;height:21px;">
<input id="BtnCancel" type=button value="Cancel" title="Cancel application login"
       class=TButton style="width:75px;height:25px;"
       style="left:215px;top:120px;" onclick="vbscript:closeApp()">
<input id="BtnOK" type=submit value="Login" title="Login to PowerTest.NET"
       class=TButton style="width:75px;height:25px;"
       style="left:135px;top:120px;">

<a		href="reguser.asp"
		class=TLinkLabel
		title="Click here if you are a new user"
		style="width:100px;left:285px;top:160px;text-align:right;">
New user signup
</a>


<div id="Form2" style="display: none; position:absolute; border:groove thin;"
                style="left:2px; top:120px; width:392px; height:120px; overflow:hidden;">
<IFRAME ID=ConfirmIframe Name=ConfirmIframe FRAMEBORDER=No FRAMESPACING=0 width=100% height=100% scrolling=no>
</IFRAME>
</div>

</div>
</form>

<script language=vbscript>
' Handles the Cancel button on click event
Sub closeApp()
	Call window.parent.frames("Header").CloseApp
End Sub


' Handless the document on load complete event
Sub document_onReadyStateChange()
	If LCase(document.readyState) = "complete" Then LoginForm.LoginUser.focus()
End Sub


' Check the form data before sending to the server
Function LoginForm_onSubmit()
	If LoginForm.LoginUser.Value = "" Then
		MsgBox "Please enter login name", vbOKOnly+vbExclamation, "PowerTests .NET"
		LoginForm_onSubmit = False
		LoginForm.LoginUser.Focus()
	ElseIf LoginForm.LoginPass.Value = "" Then
		MsgBox "Please enter password", vbOKOnly+vbExclamation, "PowerTests .NET"
		LoginForm_onSubmit = False
		LoginForm.LoginPass.Focus
	Else
		LoginForm.BtnOK.disabled = True
		LoginForm.BtnCancel.disabled = True
		LoginForm_onSubmit = True
	End If
End Function


' Diplays the DIV that contains the IFRAME where 
' the username and password form action response will be loaded
Public Sub ShowConfirm()
		LoginForm.BtnOK.style.visibility = "hidden"
		LoginForm.BtnCancel.style.visibility = "hidden"
		Form1.style.height = "246px"
		Form2.style.display = "block"
		document.frames(0).focus 
End Sub


' Hides the DIV...
Public Sub HideConfirm()
	Form1.style.height = "180px"
	Form2.style.display = "none"
	With LoginForm
		.BtnOK.style.visibility = "visible"
		.BtnOK.disabled = False
		.BtnCancel.style.visibility = "visible"
		.BtnCancel.disabled = False
		.LoginUser.Focus()
	End With
End Sub
</script>

</BODY>
</HTML>
