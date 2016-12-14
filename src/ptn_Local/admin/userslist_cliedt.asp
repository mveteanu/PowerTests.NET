<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim PersonData(5)
 Dim langOptions
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 Set RSUser = GetUserByID(Request.QueryString("userid"),cn)
 PersonData(0) = RSUser.Fields("nume").Value
 PersonData(1) = RSUser.Fields("prenume").Value
 PersonData(2) = RSUser.Fields("email").Value
 PersonData(3) = RSUser.Fields("telefon").Value
 PersonData(4) = RSUser.Fields("id_lang").Value
 RSUser.Close
 set RSUser=nothing
 
 langOptions = GetOptionsForSelect("<option value='@0' @SELECTED>@1</option>", "@0 = " & PersonData(4), "TBLanguages", Array("id", "langname"), cn)
 
 cn.Close 
 set cn=nothing
%>
<html>
<head>
  <title>Edit user</title>
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
          style="left:16px;top:8px;">
<legend>Personal info</legend>
<span id="Label1"
      class=TLabel style="width:110px;height:13px;"
      style="left:24px;top:24px;">
Last name:
</span>
<span id="Label2"
      class=TLabel style="width:110px;height:13px;"
      style="left:24px;top:56px;">
First name:
</span>
<span id="Label3"
      class=TLabel style="width:110px;height:13px;"
      style="left:24px;top:88px;">
Email:
</span>
<span id="Label4"
      class=TLabel style="width:110px;height:13px;"
      style="left:24px;top:120px;">
Phone:
</span>
<span id="Label4"
      class=TLabel style="width:110px;height:13px;"
      style="left:24px;top:152px;">
Preferred lang:
</span>
<input id="Edit1" type=text maxlength=20 value="<%=PersonData(0)%>"
       class=TEdit style="width:140px;height:21px;"
       style="left:112px;top:16px;"
       title="Last name">
<input id="Edit2" type=text maxlength=20 value="<%=PersonData(1)%>"
       class=TEdit style="width:140px;height:21px;"
       style="left:112px;top:48px;"
       title="First name">
<input id="Edit3" type=text maxlength=50 value="<%=PersonData(2)%>"
       class=TEdit style="width:140px;height:21px;"
       style="left:112px;top:80px;"
       title="Email">
<input id="Edit4" type=text maxlength=20 value="<%=PersonData(3)%>"
       class=TEdit style="width:140px;height:21px;"
       style="left:112px;top:112px;"
       title="Phone number">
<select id="cboLanguage"
		class=TComboBox
		style="left:112px;top:144px;"
		style="width:140px;height:21px;">
		<%=langOptions%>
</select>
</fieldset>
<fieldset class=TGroupBox style="width:289px;height:97px;"
          style="left:16px;top:200px;"
          title="If you do not intent to change the password, leave these fields empty">
<legend>Change password</legend>
<span id="Label5"
      class=TLabel style="width:90px;height:13px;"
      style="left:24px;top:28px;">
New password:
</span>
<span id="Label6"
      class=TLabel style="width:90px;height:13px;"
      style="left:24px;top:68px;">
Verify password:
</span>
<input id="Edit5" type=password maxlength=20 autocomplete="off" 
       class=TEditPassword style="width:140px;height:21px;"
       style="left:112px;top:20px;"
       title="New password">
<input id="Edit6" type=password maxlength=20 autocomplete="off" 
       class=TEditPassword style="width:140px;height:21px;"
       style="left:112px;top:60px;"
       title="Verify password">
</fieldset>
<input id="Button1" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:171px;top:312px;">
<input id="Button2" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:75px;top:312px;">
</div>


<script language="javascript" src="../_clientscripts/emails.js"></script>
<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Cancel
Sub Button1_onclick
  Window.close 
End Sub


' Intoarce un array cu datele continute in controale
Function GetPersData()
  Dim DatePers(6)
  
  DatePers(0) = Edit1.Value
  DatePers(1) = Edit2.Value
  DatePers(2) = Edit3.Value
  DatePers(3) = Edit4.Value
  DatePers(4) = cboLanguage.value
  DatePers(5) = Edit5.Value

  GetPersData = DatePers
End Function


'Evenimentul apare la apasarea butonului OK
Sub Button2_onclick()
  If (Edit1.Value="") or (Edit2.Value="") then
    msgbox "Last name or first name cannot be null.", vbOkOnly+vbExclamation
    Exit Sub
  ElseIf Not validEMail(Edit3.value) then
    msgbox "Invalid email address.", vbOkOnly+vbExclamation
    Exit Sub
  ElseIf Edit5.value <> Edit6.value then
    msgbox "Password verification does not match password.", vbOkOnly+vbExclamation
    Exit Sub
  End If
    
  Window.Returnvalue = GetPersData()
  Window.Close
End Sub
</script>

</body>
</html>