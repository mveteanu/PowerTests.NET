<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim cn, studinfo

 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set studinfo = GetStudInfo(Request.QueryString("studID"), cn)
 cn.Close
 set cn = nothing

 Class StudInfoRec
   Public Nume
   Public Prenume
   Public Login
   Public Telefon
   Public Email
   Public NrTesteRezolvate
   Public DataInscriereCurs
   Public DataValidareCurs
   Public BlocatLaAplicatie
   Public BlocatLaCurs
 End Class
 
 ' Intoarce un obiect de tipul StudInfoRec cu
 ' informatii despre student
 Function GetStudInfo(studID, objCon)
   Dim re
   Dim rs, myCmd
   
   Set re = New StudInfoRec
   Set myCmd = Server.CreateObject("ADODB.Command")
   Set myCmd.ActiveConnection = objCon
   myCmd.CommandText = "GetStudValByInscriereID"
   myCmd.CommandType = adCmdStoredProc
   Set rs = myCmd.Execute(,CLng(studID))
   If Not rs.EOF Then
    re.Nume    = rs.Fields("nume").Value
    re.Prenume = rs.Fields("prenume").Value
    re.Login   = rs.Fields("login").Value
    re.Telefon = rs.Fields("telefon").Value
    re.Email   = rs.Fields("email").Value
    re.NrTesteRezolvate  = rs.Fields("TesteRezolvate").Value
    re.DataInscriereCurs = rs.Fields("datainscriere").Value
    re.DataValidareCurs  = rs.Fields("datavalidare").Value
    re.BlocatLaAplicatie = rs.Fields("applocked").Value
    re.BlocatLaCurs      = rs.Fields("locked").Value 
   End If
   rs.Close
   set rs = nothing
   set myCmd = nothing
   Set GetStudInfo = re
 End Function
%>
<html>
<head>
  <title>Student summary</title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body>

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
<fieldset class=TGroupBox style="width:233px;height:145px;"
          style="left:8px;top:8px;">
<legend>Personal information</legend>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:24px;">
Last name: <%=studinfo.nume%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:48px;">
First name: <%=studinfo.prenume%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:72px;">
Login: <%=studinfo.login%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:96px;">
Phone: <%=studinfo.telefon%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:120px;">
Email: <%="<a href='mailto:" & studinfo.email & "'>" & studinfo.email & "</a>"%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:233px;height:145px;"
          style="left:248px;top:8px;">
<legend>Student in class</legend>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:24px;">
Number of testings: <%=NullToZero(studinfo.NrTesteRezolvate)%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:48px;">
Enrollment date: <%=DToSR(studinfo.DataInscriereCurs,"DD/MM/YYYY")%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:72px;">
Acceptance date: <%=DToSR(studinfo.DataValidareCurs,"DD/MM/YYYY")%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:96px;">
Locked at application level: <%=BooleanToAfirm(studinfo.BlocatLaAplicatie)%>
</span>
<span class=TLabel style="width:200px;height:13px;"
      style="left:16px;top:120px;">
Locked at course level: <%=BooleanToAfirm(studinfo.BlocatLaCurs)%>
</span>
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:207px;top:164px;">
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
