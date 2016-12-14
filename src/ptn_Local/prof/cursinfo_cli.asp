<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/cursuri.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim CursData(9)
 Dim CursID
 
 CursID = Request.QueryString("cursid")
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSC = GetCursByID(CursID,cn)
 
 CursData(0) = RSC.Fields("numecurs").Value
 CursData(1) = NrStudToString(RSC.Fields("maxstudents").Value, "")
 CursData(2) = PermisionToString(RSC.Fields("permisiiacceptare").Value)
 CursData(3) = BooleanToAfirm(RSC.Fields("curspublic").Value)
 CursData(4) = RSC.Fields("NrPb").Value
 CursData(5) = RSC.Fields("NrCateg").Value
 CursData(6) = RSC.Fields("NrTst").Value
 CursData(7) = RSC.Fields("StudentiInscrisi").Value
 CursData(8) = RSC.Fields("CereriInscriere").Value
 set RSC=nothing
 
 
' Umple listbox-ul cu studentii de la curs
Sub FillStudentsListBox(id, objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetStudValByCurs"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(id))
  do until rs.EOF
    Response.Write "<option>" & rs.Fields("numestud").Value & "</option>"
    rs.MoveNext
  loop
  rs.Close
  Set rs = nothing
  Set myCmd = Nothing
End Sub


' Umple listbox-ul cu cererile pentru curs
Sub FillCereriListBox(id, objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetCereriByCurs"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(id))
  do until rs.EOF
    Response.Write "<option>" & rs.Fields("numestud").Value & "</option>"
    rs.MoveNext
  loop
  rs.Close
  Set rs = nothing
  Set myCmd = Nothing
End Sub
%>
<html>
<head>
  <title>Course summary</title>
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
<fieldset class=TGroupBox style="width:225px;height:129px;"
          style="left:8px;top:8px;">
<legend>Course properties</legend>
<span id="Label1"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:24px;">
Course name: <%=CursData(0)%>
</span>
<span id="Label2"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:48px;">
Maximum students: <%=CursData(1)%>
</span>
<span id="Label3"
      class=TLabel style="width:210px;height:13px;"
      style="left:8px;top:72px; ">
Enrolment policy: <%=CursData(2)%>
</span>
<span id="Label4"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:96px;">
Course is public: <%=CursData(3)%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:225px;height:144px;"
          style="left:8px;top:144px;">
<legend>Current couse state</legend>
<span id="Label5"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:24px;">
Questions: <%=CursData(4)%>
</span>
<span id="Label6"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:48px;">
Question categories: <%=CursData(5)%>
</span>
<span id="Label7"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:72px;">
Tests: <%=CursData(6)%>
</span>
<span id="Label8"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:96px;">
Enrolled students: <%=CursData(7)%>
</span>
<span id="Label9"
      class=TLabel style="width:200px;height:13px;"
      style="left:8px;top:120px;">
Student requests: <%=CursData(8)%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:200px;height:280px;"
          style="left:240px;top:8px;">
<legend>Enrolled students</legend>
<select id="ListBox1" 
        size=2 style="background-color:buttonface;"
        class=TListBox style="width:180px;height:260px;"
        style="left:8px;top:16px;">
<%FillStudentsListBox CursID, cn%>
</select>
</fieldset>
<fieldset class=TGroupBox style="width:200px;height:280px;"
          style="left:446px;top:8px;">
<legend>Student requests to attend couse</legend>
<select id="ListBox2" 
        size=2 style="background-color:buttonface;"
        class=TListBox style="width:180px;height:260px;"
        style="left:8px;top:16px;">
<%FillCereriListBox CursID, cn%>
</select>
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:292px;top:298px;">
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
 cn.Close 
 set cn=nothing
%>