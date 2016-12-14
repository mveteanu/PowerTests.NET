<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/tests.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 Dim tstinfo, tstsustineri
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 set tstinfo = New PTNTestDefinition
 tstinfo.LoadTestDefinition CLng(Request.QueryString("TestID")), cn
 tstsustineri = GetTestNrSustineriByUser(Request.QueryString("TestID"), Session("UserID"), cn)
 cn.Close 
 set cn = nothing
%>
<html>
<head>
  <title>Test info</title>
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
<fieldset class=TGroupBox style="width:313px;height:193px;"
          style="left:8px;top:8px;">
<legend>Test details</legend>
<span class=TLabel style="width:64px;height:13px;"
      style="left:8px;top:32px;">
Test name:
</span>
<input READONLY type=text class=TEdit style="left:74px;top:24px;"
       style="width:231px;height:21px;" value="<%=tstinfo.name%>">
<span class=TLabel style="width:52px;height:13px;"
      style="left:8px;top:64px;">
Comments:
</span>
<textarea READONLY class=TEdit style="width:297px;height:105px;"
          style="overflow:auto;left:6px;top:80px;">
<%=tstinfo.comments%>          
</textarea>
</fieldset>
<fieldset class=TGroupBox style="width:185px;height:193px;"
          style="left:328px;top:8px;">
<legend>Test properties</legend>
<span class=TLabel style="width:26px;height:13px;"
      style="left:8px;top:32px;">
Time:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:8px;top:80px;">
Allowed solvings:
</span>
<span class=TLabel style="width:100px;height:13px;"
      style="left:8px;top:136px;">
Solvings:
</span>
<input READONLY type=text class=TEdit style="left:48px;top:24px;"
       style="width:129px;height:21px;" value="<%=tstinfo.time%>">
<input READONLY type=text class=TEdit style="left:8px;top:96px;"
       style="width:169px;height:21px;" value="<%=CompareToString2(tstinfo.MaxSustineri,0, "Unlimited")%>">
<input READONLY type=text class=TEdit style="left:8px;top:152px;"
       style="width:169px;height:21px;" value="<%=tstsustineri%>">
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:224px;top:212px;">
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
<%
set tstinfo = nothing
%>