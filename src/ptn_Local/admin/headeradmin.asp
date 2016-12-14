<%@ Language=VBScript %>
<html>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body style="background-color: buttonface" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0"
       unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<table unselectable="on" class="window" cellspacing=0 cellpadding=0 WIDTH="100%" HEIGHT=54><tr>
<td valign=center align=left>
	<div unselectable="on" id="butCereri" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Manage new user sign-up requests">
		<table unselectable="on" style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile1.png"><br>Manage<br>requests&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div unselectable="on" id="butUseri" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Manage application users">
		<table unselectable="on" style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile2.png"><br>Manage<br>users&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div unselectable="on" id="butSettings" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Application settings...">
		<table unselectable="on" style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/tools.png"><br>Application<BR>settings&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
</td>
<td valign=center align=right>
	<div unselectable="on" id="butExit" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Close application">
		<table unselectable="on" style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/dooropen.png"><br>Exit</font>
		</td></tr></table>
	</div>
</td>
</tr></table>


<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language=vbscript>
Dim mymenu1

MenuItems1 = Array("Summary of requests","<HR>","Administrator requests","Professor requests","Student requests")
MenuItems2 = Array("Summary of users","<HR>","Manage administrators","Manage professors","Manage students")
MenuItems3 = Array("Manage system languages", "Translate texts", "<HR>", "New user sign-up policy","<HR>","Personal account settings")

Sub butCereri_onclick
	set mymenu1 = showmenu(0 , 52 ,150,"handlemenuclick1",MenuItems1)
End Sub

Sub butUseri_onclick
	set mymenu1 = showmenu(90 ,52,150,"handlemenuclick1",MenuItems2)
End Sub

Sub butSettings_onclick
	set mymenu1 = showmenu(176 ,52,190,"handlemenuclick1",MenuItems3)
End Sub

Sub butExit_onclick
	CloseApp
End Sub

Sub handlemenuclick1(html)
	If html="<HR>" Or html="" Then Exit Sub
	mymenu1.Hide
	Set mymenu1 = Nothing

	Select Case html
		Case MenuItems1(0) window.parent.frames("Main").location = "cererisumar_cli.asp"
		Case MenuItems1(2) window.parent.frames("Main").location = "cererilista_cli.asp?tipuser=A"
		Case MenuItems1(3) window.parent.frames("Main").location = "cererilista_cli.asp?tipuser=P"
		Case MenuItems1(4) window.parent.frames("Main").location = "cererilista_cli.asp?tipuser=S"

		Case MenuItems2(0) window.parent.frames("Main").location = "userisumar_cli.asp"
		Case MenuItems2(2) window.parent.frames("Main").location = "userslist_cli.asp?tipuser=A"
		Case MenuItems2(3) window.parent.frames("Main").location = "userslist_cli.asp?tipuser=P"
		Case MenuItems2(4) window.parent.frames("Main").location = "userslist_cli.asp?tipuser=S"

		Case MenuItems3(0) window.parent.frames("Main").location = "languageslist_cli.asp"
		Case MenuItems3(1) window.parent.frames("Main").location = "languagetexts_cli.asp"
		Case MenuItems3(3) window.parent.frames("Main").location = "modifpolitica_cli.asp"
		Case MenuItems3(5) window.parent.frames("Main").location = "modifyourself_cli.asp"
	End Select
End Sub

Sub closecmd_onclick
	window.event.returnValue = false
	CloseApp
End Sub

Sub CloseApp
	If MsgBox("Are you sure you want to quit?",vbYesNo+vbQuestion,"Confirm") = vbYes Then window.parent.close
End Sub
</script>

</body>
</html>

