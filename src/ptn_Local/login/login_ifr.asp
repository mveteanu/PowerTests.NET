<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
ProcessInputData


' Process the form data (login and password) and displays
' a message: authentification OK, inexistent user, locked account, etc.
Sub ProcessInputData
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("DSN")

	Set RSUser = GetUser(Request.Form("LoginUser"), Request.Form("LoginPass"), cn)
	If RSUser.EOF And RSUser.BOF Then
		PassRetryNo = Session("PassRetryNo") + 1
		If PassRetryNo < 3 Then
			showError "User cannot be found or password incorrect!", True
		Else
			showError "Connection to PTN was denied. Too many retries.!", False
		End If
		Session("PassRetryNo") = PassRetryNo
	ElseIf IsNull(RSUser.Fields("datavalidare").Value) then
		showError "Account is not active!<BR>Please have patience. An administrator will review your submission soon.", False
	ElseIf RSUser.Fields("locked").Value = true then
		showError "Your account is temporary disabled!<BR>Please contact an administrator for more information.", False
	Else
		showConfirm RSUser.Fields("tipuser").Value
		Session("UserID") = RSUser.Fields("id_user").Value
	End If
  
	RSUser.Close
	Set RSUser = Nothing
	cn.Close 
	Set cn = Nothing
End Sub


' Displays the error message if the user cannot be validated. 
' If bRetry = true then the retry button will be displayed also
' A user cannot access the application in the following situations:
'	- incorrect username or password
'	- the account was not yet activated by an administrator
'	- the account is temporary locked
Sub showError(strError, bRetry)
%>
<html>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" style="background-color: ButtonFace" style="behavior:url('../_clientscripts/application.htc');">

<span unselectable="on" class=TLabel style="width:70px;height:13px;font-weight:bold;"
      style="left:161px;top:16px;">
Login error!
</span>
<span unselectable="on" class=TLabel style="width:373px;height:13px;font-weight:bold;"
      style="left:10px;top:48px;text-align:center;color:indianred;">
<%=strError%>
</span>
<%If bRetry Then%>
<input id="BtnCancel" type=button value="Cancel" title="Cancel login to application"
       class=TButton style="width:75px;height:25px;"
       style="left:94px;top:80px;">
<input id="BtnRetry" type=button value="Retry" title="Retry login"
       class=TButton style="width:75px;height:25px;"
       style="left:223px;top:80px;">
<%Else%>
<input id="BtnCancel" type=button value="Cancel" title="Cancel login to application"
       class=TButton style="width:75px;height:25px;"
       style="left:159px;top:80px;">
<%End If%>

<script Language="VBScript">
Sub window_onLoad()
	If LCase(document.readyState) = "complete" Then 
		window.Parent.ShowConfirm
		<%If bRetry Then Response.Write "btnRetry.Focus()" Else Response.Write "btnCancel.Focus()"%>
	End If 
End Sub

Sub btnCancel_onClick
	Call window.parent.parent.frames("Header").CloseApp
End Sub

Sub btnRetry_onClick
	window.Parent.HideConfirm 
End Sub
</script>

</body>
</html>
<%
End Sub


' Displayes a confirmation message if the user was successfully validated
' and allows the user to enter in his section specified by UserType = "A","P","S"
Sub showConfirm(UserType)
	Dim UserTypeFull
	Dim PageToGoH

	Select Case UserType
		Case "A" UserTypeFull = "Administrator"
			PageToGoH = "../admin/headeradmin.asp"
		Case "P" UserTypeFull = "Professor"
			PageToGoH = "../prof/headerprof.asp"
		Case "S" UserTypeFull = "Student"
		PageToGoH = "../stud/headerstud.asp"
	End Select
%>
<html>
<head>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body unselectable="on" style="background-color: ButtonFace" style="behavior:url('../_clientscripts/application.htc');">

<span unselectable="on" class=TLabel style="width:119px;height:13px;font-weight:bold;"
      style="left:148px;top:16px;">
Access granted
</span>
<span unselectable="on" class=TLabel style="width:240px;height:13px;font-weight:bold;"
      style="left:77px;top:48px;text-align:center;">
Enter application as <%=UserTypeFull%>
</span>
<input id="BtnOK" type=button value="OK" title="Intra in aplicatie"
       class=TButton style="width:75px;height:25px;"
       style="left:159px;top:80px;">

<script Language="VBScript">
Sub window_onLoad()
	window.Parent.ShowConfirm 
	btnOK.focus
End Sub

Sub btnOK_onClick
	btnOK.disabled = True
	window.Parent.parent.frames("Header").location = "<%=PageToGoH%>"
	window.Parent.parent.frames("Main").location = "middle.asp"
End Sub
</script>

</body>
</html>
<%
End Sub
%>
