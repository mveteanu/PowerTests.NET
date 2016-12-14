<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/Utils.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim PersonData(6)
 Dim UserID
 
 UserID = Request.QueryString("userid")
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSUser = GetUserByID(UserID,cn)
 Set RSUser2 = GetValidatedUsers("S",cn)
 RSUser2.Filter = "id_user=" & UserID
 
 PersonData(0) = RSUser.Fields("nume").Value
 PersonData(1) = RSUser.Fields("prenume").Value
 PersonData(2) = RSUser.Fields("email").Value
 PersonData(3) = RSUser.Fields("telefon").Value
 PersonData(4) = DToSR(RSUser.Fields("datavalidare").Value, "DD/MM/YYYY")
 PersonData(5) = NullToZero(RSUser2.Fields("CursuriInscris").Value)
 RSUser.Close
 RSUser2.Close
 set RSUser=nothing
 set RSUser2=nothing

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
<fieldset class=TGroupBox style="width:215px;height:193px;"
          style="left:10px;top:8px;">
<legend>Personal info</legend>
<span id="Label1"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:24px;">
Accout type: STUDENT
</span>
<span id="Label2"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:48px;">
Last name: <%=PersonData(0)%>
</span>
<span id="Label3"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:72px;">
First name: <%=PersonData(1)%>
</span>
<span id="Label4"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:96px;">
Email: <a href="mailto:<%=PersonData(2)%>"><%=PersonData(2)%></a>
</span>
<span id="Label5"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:120px;">
Phone: <%=PersonData(3)%>
</span>
<span id="Label6"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:144px;">
Acceptance date: <%=PersonData(4)%>
</span>
<span id="Label7"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:168px;">
Courses: <%=PersonData(5)%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:345px;height:193px;"
          style="left:232px;top:8px;">
<legend>Courses</legend>
<%CreateTableControl 8, 16, 161, Array("Course", "Professor"), Array(150,179), 0, "usersinfos_dat.asp?iduser="& Cstr(UserID) , false, "MyTableDet"%>
</fieldset>

<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:253px;top:216px;">
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
