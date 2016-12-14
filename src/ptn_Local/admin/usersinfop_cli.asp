<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/Utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim PersonData(7)
 Dim UserID
 
 UserID = Request.QueryString("userid")
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSUser = GetUserByID(UserID,cn)
 Set RSUser2 = GetValidatedUsers("P",cn)
 RSUser2.Filter = "id_user=" & UserID
 
 PersonData(0) = RSUser.Fields("nume").Value
 PersonData(1) = RSUser.Fields("prenume").Value
 PersonData(2) = RSUser.Fields("email").Value
 PersonData(3) = RSUser.Fields("telefon").Value
 PersonData(4) = DToSR(RSUser.Fields("datavalidare").Value, "DD/MM/YYYY")
 PersonData(5) = NullToZero(RSUser2.Fields("NumarCursuri").Value)
 PersonData(6) = NullToZero(RSUser2.Fields("NumarStudenti").Value)
 RSUser.Close
 RSUser2.Close
 set RSUser=nothing
 set RSUser2=nothing

' Umple primul listbox cu cursurile profesorului
Sub FillCursuriListBox(id, objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetCursuriByProf"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(id))
  do until rs.EOF
    Response.Write "<option>" & rs.Fields("numecurs").Value & " ("& rs.Fields("StudentiInscrisi").Value &")</option>"
    rs.MoveNext
  loop
  rs.Close
  Set rs = nothing
  Set myCmd = Nothing
End Sub
 
' Umple al doilea listbox cu studentii profesorului
Sub FillStudentsListBox(id, objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetStudentsByProf"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(id))
  do until rs.EOF
    Response.Write "<option>" & rs.Fields("numestud").Value & " ("& rs.Fields("numarinscrierilaprof").Value &")</option>"
    rs.MoveNext
  loop
  rs.Close
  Set rs = nothing
  Set myCmd = Nothing
End Sub
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
<fieldset class=TGroupBox style="width:215px;height:217px;"
          style="left:10px;top:8px;">
<legend>Personal info</legend>
<span id="Label1"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:24px;">
Account type: PROFESSOR
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
Number of courses: <%=PersonData(5)%>
</span>
<span id="Label8"
      class=TLabel style="width:180px;height:13px;"
      style="left:16px;top:192px;">
Number of students: <%=PersonData(6)%>
</span>
</fieldset>
<fieldset class=TGroupBox style="width:345px;height:217px;"
          style="left:232px;top:8px;">
<legend>Details</legend>
<span id="Label9"
      class=TLabel style="width:155px;height:13px;"
      style="left:8px;top:16px;">
Courses ( Students )
</span>
<span id="Label10"
      class=TLabel style="width:97px;height:13px;"
      style="left:180px;top:16px;">
Students ( Courses )
</span>
<select id="ListBox1" 
        size=2 style="background-color:buttonface;"
        class=TListBox style="width:153px;height:177px;"
        style="left:8px;top:32px;">
<%FillCursuriListBox UserID, cn%>
</select>
<select id="ListBox2" 
        size=2 style="background-color:buttonface;"
        class=TListBox style="width:153px;height:177px;"
        style="left:180px;top:32px;">
<%FillStudentsListBox UserID, cn%>
</select>
</fieldset>
<input id="Button1" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:255px;top:240px;">
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