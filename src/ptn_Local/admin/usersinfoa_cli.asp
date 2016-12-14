<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/Utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim PersonData(5)
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSUser = GetUserByID(Request.QueryString("userid"),cn)
 PersonData(0) = RSUser.Fields("nume").Value
 PersonData(1) = RSUser.Fields("prenume").Value
 PersonData(2) = RSUser.Fields("email").Value
 PersonData(3) = RSUser.Fields("telefon").Value
 PersonData(4) = DToSR(RSUser.Fields("datavalidare").Value, "DD/MM/YYYY")
 RSUser.Close
 set RSUser=nothing
 cn.Close 
 set cn=nothing
%>
<html>
<head>
  <title>Account summary</title>
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
<fieldset class=TGroupBox style="width:289px;height:185px;"
          style="left:10px;top:8px;">
<legend>Personal info</legend>
<span id="Label1"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:24px;">
Account type: ADMINISTRATOR
</span>
<span id="Label2"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:48px;">
Last name: <%=PersonData(0)%>
</span>
<span id="Label3"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:72px;">
First name: <%=PersonData(1)%>
</span>
<span id="Label4"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:96px;">
Email: <a href="mailto:<%=PersonData(2)%>"><%=PersonData(2)%></a>
</span>
<span id="Label5"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:120px;">
Phone: <%=PersonData(3)%>
</span>
<span id="Label6"
      class=TLabel style="width:250px;height:13px;"
      style="left:16px;top:144px;">
Acceptance date: <%=PersonData(4)%>
</span>
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:118px;top:210px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button1_onclick
  Window.close 
End Sub
</script>

</body>
</html>