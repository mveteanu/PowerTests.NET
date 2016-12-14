<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/cursuri.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 Dim CursData(7)
 Dim CursID
 CursID = Request.QueryString("CursID")
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSC = GetCursByID(CursID,cn)
 If not (RSC.BOF and RSC.EOF) Then
	CursData(0) = RSC.Fields("numecurs").Value
	CursData(1) = PermisionToString(RSC.Fields("permisiiacceptare").Value)
	CursData(2) = NrStudToString(RSC.Fields("maxstudents").Value, "")
	CursData(3) = RSC.Fields("StudentiInscrisi").Value
	CursData(4) = RSC.Fields("CereriInscriere").Value
	CursData(5) = RSC.Fields("profname").Value 
	CursData(6) = "<a href='mailto:" & RSC.Fields("email").Value & "'>" & RSC.Fields("email").Value & "</a>"
 End If
 set RSC=nothing
 cn.Close
 set cn = nothing
%>
<html>
<head>
  <title>Course info</title>
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
<fieldset class=TGroupBox style="width:425px;height:73px;"
          style="left:8px;top:8px;">
<legend>Professor info</legend>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:24px;">
Professor name: <%=CursData(5)%>
</span>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:48px;">
Email: <%=CursData(6)%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:425px;height:145px;"
          style="left:8px;top:88px;">
<legend>Course info</legend>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:24px;">
Course name: <%=CursData(0)%>
</span>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:48px;">
Enrollment policy: <%=CursData(1)%>
</span>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:72px;">
Maximum number of students: <%=CursData(2)%>
</span>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:96px;">
Enrolled students: <%=CursData(3)%>
</span>
<span class=TLabel style="width:400px;height:13px;"
      style="left:16px;top:120px;">
Enrollment requests: <%=CursData(4)%>
</span>
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:184px;top:246px;">
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
