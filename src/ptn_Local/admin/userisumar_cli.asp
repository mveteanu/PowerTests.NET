<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
  Set cn=Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  SumarUseri = GetSumarUsers(cn)
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
<legend>Summary of system users</legend>
<span id="Label1"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:20px;">
Administrators: <%=SumarUseri(3)%>
</span>
<span id="Label2"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:44px;">
Professors: <%=SumarUseri(4)%>
</span>
<span id="Label3"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:68px;">
Students: <%=SumarUseri(5)%>
</span>

<span id="Label4"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:100px;">
Locked administrators: <%=SumarUseri(6)%>
</span>
<span id="Label5"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:124px;">
Locked professors: <%=SumarUseri(7)%>
</span>
<span id="Label6"
      class=TLabel style="width:120px;height:13px;"
      style="left:16px;top:148px;">
Locked students: <%=SumarUseri(8)%>
</span>

<img width=250 height=180 border=0 class="TImage"
     style="left:167px;top:13px;"
     src="userisumar_img.asp?A=<%=SumarUseri(3)%>&P=<%=SumarUseri(4)%>&S=<%=SumarUseri(5)%>&AB=<%=SumarUseri(6)%>&PB=<%=SumarUseri(7)%>&SB=<%=SumarUseri(8)%>">
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:90px;height:25px;"
       style="left:185px;top:230px;">
</div>


<script language=vbscript>
' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub window_onload
  If <%=SumarUseri(0)%>=-1 then
    msgbox "Error obtaining data.",vbOkOnly+vbCritical
    HideAllDivs
    Exit Sub
  End If
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub



' Ascunde toate div-urile, adica cel cu fereastra si cel cu mesajul de wait
Sub HideAllDivs
  Form1.style.visibility = "hidden"
  WaitforForm.style.visibility = "hidden"
End Sub


' Trateaza evenimentul care apare la apasarea butonului Close
Sub Button1_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
