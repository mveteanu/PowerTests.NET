<%@ Language=VBScript %>
<%
 Response.Buffer = True
 Response.Expires = -1
%>
<html>
<head>
  <title>Preview questions before print</title>
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

<iframe id=rezframe name=rezframe src="pbviewprn_ifr.asp?<%=Request.QueryString%>"
        style="position:absolute;left:8px;top:8px;width:698px;height:433px;">
</iframe>

<input DISABLED id="Button1" type=button value="Hide texts" title="Hide/Show questions text"
       class=TButton style="width:120px;height:25px;"
       style="left:36px;top:456px;">
<input DISABLED id="Button2" type=button value="Hide answers" title="Hide/Show questions answer"
       class=TButton style="width:120px;height:25px;"
       style="left:209px;top:456px;">
<input DISABLED id="Button3" type=button value="Print" title="Send page to printer"
       class=TButton style="width:120px;height:25px;"
       style="left:383px;top:456px;">
<input id="Button4" type=button value="Close" title="Close form"
       class=TButton style="width:120px;height:25px;"
       style="left:556px;top:456px;">
</div>


<div id="WaitforFormIfr" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:8px;top:8px;width:698px;height:433px;">
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

' Schimba starea activ/inactiv a butoanelor
Sub ActivateButtons(btnsstate)
 Dim i
 
 for i=1 to 3
   Form1.all("Button"&CStr(i)).disabled = not btnsstate
 next  
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

' Evenimentul care apare la apasarea butonului ce comuta 
' intre afisarea/ascunderea enunturilor
Sub Button1_onclick
 Dim oldstyle
 oldstyle = LCase(Window.Frames("rezframe").tblEnunturi.style.display)
 
 If oldstyle = "none" then
   Window.Frames("rezframe").tblEnunturi.style.display = ""
   Button1.value = "Hide texts"
 Else
   Window.Frames("rezframe").tblEnunturi.style.display = "none"
   Button1.value = "Show texts"
 End If  
End Sub


' Evenimentul care apare la apasarea butonului ce comuta 
' intre afisarea/ascunderea raspunsurilor
Sub Button2_onclick
 Dim oldstyle
 oldstyle = LCase(Window.Frames("rezframe").tblRaspunsuri.style.display)
 
 If oldstyle = "none" then
   Window.Frames("rezframe").tblRaspunsuri.style.display = ""
   Button2.value = "Hide answers"
 Else
   Window.Frames("rezframe").tblRaspunsuri.style.display = "none"
   Button2.value = "Show answers"
 End If  
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
