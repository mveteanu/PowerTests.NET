<%@ Language=VBScript %>
<%
If Request.QueryString.Count <> 0 Then
	Call PrintContactServerOK
Else  
	Call PrintNormalFooter
End If


' Subroutine is called automatically 
' at constant time intervals
' to avoid session expiration
Sub PrintContactServerOK
%>
<body>
<script language=vbscript>
	Call Window.parent.ContactWebServerOK
</script>
</body>
<%
End Sub 


' Displays the footer of application pages
Sub PrintNormalFooter
%>
<html>
<head>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body style="background-color: buttonface" topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0
      unselectable="on" style="behavior:url('../_clientscripts/application.htc');">
      
<table border=0 style='border:outset thin;' width=100% height=100% cellspacing=0 cellpadding=2>
<tr>
 <td align="left" valign="middle"><span unselectable="on" id=LabelStatus>&nbsp;</span></td>
 <td style='border:inset thin;' align="center" valign="middle" width="50"><span unselectable="on" style="cursor:default;" id=LabelTime title="Client system time">00:00</span></td>
</tr>
</table>

<IFRAME style="display:none;"></IFRAME>

<script language=vbscript>
Dim MinuteScurse
Dim ContactServerInterval


Sub window_onload
	ContactServerInterval = 10 ' the web server is contacted at each 10 minutes
	window.setInterval "UpdateClock", 60000
	Call UpdateClock
End Sub


Sub UpdateClock
	LabelTime.innerText = FormatDateTime(Now(),vbShortTime)
 
	MinuteScurse = MinuteScurse + 1
	If (MinuteScurse mod ContactServerInterval) = 0 Then ContactWebServer
End Sub


Sub LabelTime_ondblclick
	msgbox "PowerTest .NET" & vbCrLf & vbCrLf & "© VMA soft 2001" & vbCrLf & "http://vmasoft.hypermart.net", vbOkOnly+vbInformation, "PTN Info"
End Sub


' The web server is contacted at constant time intervals
' in order to prevent user session expiration !!!!
' Is not executed in case of opened modal windows :-(
Sub ContactWebServer
	LabelStatus.innerHTML = "WebServer autocontacting..."
	document.frames(0).location.href = "../login/footer.asp?AvoidSessionExpire=yes"
End Sub


Sub ContactWebServerOK
	LabelStatus.innerHTML = "&nbsp;"
End Sub
</script>

</body>
</html>
<%
End Sub
%>