<%@ Language=VBScript %>
<%
 Response.Buffer = True
 Response.Expires = -1
%>
<html>
<head>
  <title>Test results</title>
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

<iframe id=rezframe name=rezframe src="studtstviewrez_ifr.asp?tip=1&<%=Request.QueryString%>"
        style="position:absolute;left:8px;top:8px;width:596px;height:433px;">
</iframe>

<input DISABLED id="Button1" type=button value="Group by test" title="Group displayed results by test"
       class=TButton style="width:120px;height:25px;"
       style="left:17px;top:456px;">
<input DISABLED id="Button2" type=button value="Group by score" title="Group displayed results by score"
       class=TButton style="width:120px;height:25px;"
       style="left:169px;top:456px;">
<input DISABLED id="Button3" type=button value="Print" title="Print results"
       class=TButton style="width:120px;height:25px;"
       style="left:321px;top:456px;">
<input id="Button4" type=button value="Close" title="Close form"
       class=TButton style="width:120px;height:25px;"
       style="left:473px;top:456px;">
</div>


<div id="WaitforFormIfr" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:8px;top:8px;width:596px;height:433px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Subrutina trebuie apelata din pagina ce se incarca in IFRAME
Sub HandleIframeLoading(putwait)
 If putwait then
   WaitforFormIfr.style.visibility = "visible"
   ActivateButtons false
 Else
   WaitforFormIfr.style.visibility = "hidden"
   ActivateButtons true
 End If  
End Sub

' Stabileste starea de activare a butoanelor
Sub ActivateButtons(state)
 For i = 1 to 3
  document.all("Button" & CStr(i)).disabled = not state
 Next
End Sub

' Evenimentul apare la apasarea unuia din butoanele de comutare a iframe-ului
Sub Button1_onclick
 HandleIframeLoading true
 Window.Frames("rezframe").location.href = "studtstviewrez_ifr.asp?tip=1&<%=Request.QueryString%>"
End Sub

' Evenimentul apare la apasarea unuia din butoanele de comutare a iframe-ului
Sub Button2_onclick
 HandleIframeLoading true
 Window.Frames("rezframe").location.href = "studtstviewrez_ifr.asp?tip=2&<%=Request.QueryString%>"
End Sub

' Evenimentul apare la apasarea butonului Print
Sub Button3_onclick
  Window.Frames("rezframe").focus
  Window.print 
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button4_onclick
  Window.close 
End Sub
</script>

</body>
</html>
