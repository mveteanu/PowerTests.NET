<%@ Language=VBScript %>
<%
 Response.Buffer = True
 Response.Expires = -1
%>
<html>
<head>
  <title>Student results</title>
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

<iframe id=rezframe name=rezframe src="studclasament1_ifr.asp?<%=Request.QueryString%>"
        style="position:absolute;left:8px;top:8px;width:596px;height:433px;">
</iframe>

<input DISABLED id="Button1" type=button value="Print" title="Print results"
       class=TButton style="width:120px;height:25px;"
       style="left:166px;top:456px;">
<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:120px;height:25px;"
       style="left:326px;top:456px;">
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
   Button1.disabled = true
 Else
   WaitforFormIfr.style.visibility = "hidden"
   Button1.disabled = false
 End If  
End Sub

' Evenimentul apare la apasarea butonului Print
Sub Button1_onclick
  Window.Frames("rezframe").focus
  Window.print 
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>
