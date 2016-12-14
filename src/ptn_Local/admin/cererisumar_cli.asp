<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
  Set cn=Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  SumarCereri = GetSumarUsers(cn)
  cn.Close 
  Set cn=nothing
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:460px;height:270px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:460px;height:270px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<fieldset class=TGroupBox style="width:433px;height:209px;"
          style="left:9px;top:8px;">
<legend>Summary of requests</legend>
<span id="Label1"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:20px;">
Administrator requests: <%=SumarCereri(0)%>
</span>
<span id="Label2"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:44px;">
Professor requests: <%=SumarCereri(1)%>
</span>
<span id="Label3"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:68px;">
Student requests: <%=SumarCereri(2)%>
</span>
<img width=250 height=180 border=0 class="TImage"
     style="left:167px;top:13px;"
     src="cererisumar_img.asp?A=<%=SumarCereri(0)%>&P=<%=SumarCereri(1)%>&S=<%=SumarCereri(2)%>">
</fieldset>
<input id="Button1" type=button value="Accept all" title="Grant access to all users"
       class=TButton style="width:90px;height:25px;"
       style="left:54px;top:230px;">
<input id="Button2" type=button value="Deny all" title="Deny all requests"
       class=TButton style="width:90px;height:25px;"
       style="left:180px;top:230px;">
<input id="Button3" type=button value="Close" title="Close form"
       class=TButton style="width:90px;height:25px;"
       style="left:306px;top:230px;">
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="cererisumar_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>


<script language=vbscript>
' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub window_onload
  If <%=SumarCereri(0)%>=-1 then
    msgbox "Error obtaining summary of requests",vbOkOnly+vbCritical
    HideAllDivs
    Exit Sub
  End If
  If <%=SumarCereri(0)%>+<%=SumarCereri(1)%>+<%=SumarCereri(2)%>=0 then ActivateButtons false, false, true
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


' Ascunde toate div-urile, adica cel cu fereastra si cel cu mesajul de wait
Sub HideAllDivs
  Form1.style.visibility = "hidden"
  WaitforForm.style.visibility = "hidden"
End Sub


' Schimba starea activ/inactiv a celor trei butoane
' Daca b1, b2, b3 = true butoanele sunt active si invers
Sub ActivateButtons(b1, b2, b3)
  Button1.disabled = not b1
  Button2.disabled = not b2
  Button3.disabled = not b3
End Sub


' Trateaza evenimentul care apare la apasarea butonului Accept all
Sub Button1_OnClick
 if msgbox("Are you sure you want to accept all requests without viewing detailed information?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub
 
 FormularH.SelectAction.value = "acceptall"
 ActivateButtons false, false, false
 FormularH.submit 
End Sub


' Trateaza evenimentul care apare la apasarea butonului Deny all
Sub Button2_OnClick
 if msgbox("Are you sure you want to deny all requests without viewing detailed information?",vbYesNo+vbQuestion,"Confirm") = vbNo then Exit Sub
 
 FormularH.SelectAction.value = "denyall"
 ActivateButtons false, false, false
 FormularH.submit 
End Sub


' Trateaza evenimentul care apare la apasarea butonului Close
Sub Button3_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
