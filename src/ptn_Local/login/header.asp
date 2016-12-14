<%@ Language=VBScript %>
<html>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body style="background-color: buttonface" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0"
      unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<table unselectable="on" class="window" cellspacing=0 cellpadding=0 WIDTH="100%" HEIGHT=54>
<tr>
<td valign=center align=left>
    <img src="../images/ptnlogo.png" width="105" height="37" border="0" align=middle>
</td>
<td valign=center align=right>
	<div id="closecmd" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Close PowerTest.NET">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/dooropen.png"><br>Exit</font>
		</td></tr></table>
	</div>
</td>
</tr></table>

<script language="vbscript">
Sub closecmd_onclick
	window.event.returnValue = false
	CloseApp
End Sub

Sub CloseApp
	If MsgBox("Are you sure you want to quit?", vbYesNo+vbQuestion, "Please confirm") = vbYes Then 
		window.parent.close
	End If  
End Sub
</script>

</body>
</html>

