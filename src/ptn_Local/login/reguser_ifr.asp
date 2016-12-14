<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/settings.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<%
Response.Buffer = True
Response.Expires = -1
ProcessInputData

Sub ProcessInputData
	On Error Resume Next
	Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.Open Application("DSN")
	objCon.BeginTrans 
	Call ProcessForm(objCon)
	If Err.number <> 0 Then
		objCon.RollbackTrans 
		Call PrintFormResponse("DB Error (" & Err.number & "): " & Err.Description, "indianred", true)
		Err.Clear 
	Else
		objCon.CommitTrans 
	End If
	objCon.Close
	Set objCon = Nothing
End Sub


' Function returns if the user with specified login name exists in the database
Function ExistUser(login, cn)
	Set rs = Server.CreateObject("ADODB.Recordset")
	cn.GetUsersNr CStr(login), rs
	ExistUser = CBool(rs.Fields("NrUsers").Value > 0)
	rs.Close 
	Set rs = Nothing
End Function


' Process the form submited by the user and enters the data in database
Sub ProcessForm(cn)
	Const msgExistUser       = "Another user with the same login allready exists in the system."
	Const msgRespingere      = "Your request for an application account was denied by the system. The administrator blocked new user sign-up."
	Const msgAcceptImediat   = "Sign-up successfull. Please use your login information to access the application."
	Const msgAcceptAsteptare = "Your request was successfully received into the system. Please wait for an administrator to unlock your account."

	Dim UserPolicy
	UserPolicy = GetUserPolicy(Request.Form("tipuser"),cn)

	If UserPolicy = 2 Then 
		Call PrintFormResponse(msgRespingere, "indianred", false)         ' Cererea este respinsa automat
	Else
		If ExistUser(Request.Form("login"), cn) Then
			Call PrintFormResponse(msgExistUser, "indianred", true)
		Else
			Msgs = Array(msgAcceptImediat, msgAcceptAsteptare)
			Call EnterFormInDataBase(CBool(UserPolicy = 0), cn)
			Call PrintFormResponse(Msgs(UserPolicy), "black", false)
		End If
	End If  
End Sub 'Process Form


' Inserts the form data in database
' If activeazacont = true then the account will be activate automatically (nice for students)
Sub EnterFormInDataBase(activeazacont, cn)
	Set rs1=Server.CreateObject("ADODB.Recordset")
	Set rs2=Server.CreateObject("ADODB.Recordset")
	rs1.Open "TBPersons", cn, adOpenDynamic, adLockOptimistic, adCmdTable
	rs2.Open "TBUsers", cn, adOpenDynamic, adLockOptimistic, adCmdTable
	rs1.AddNew 
	rs2.AddNew 
	rs1.Fields("nume").Value		= RemoveTags(Request.Form("nume").Item, 20)
	rs1.Fields("prenume").Value		= RemoveTags(Request.Form("prenume").Item, 20)
	rs1.Fields("email").Value		= RemoveTags(Request.Form("email").Item, 50)
	rs1.Fields("telefon").Value		= RemoveTags(Request.Form("telefon").Item, 20)
	rs2.Fields("tipuser").Value		= Request.Form("tipuser").Item 
	rs2.Fields("id_person").Value	= rs1.Fields("id_person").Value
	rs2.Fields("login").Value		= RemoveTags(Request.Form("login").Item, 20)
	rs2.Fields("pass").Value		= Request.Form("pass").Item
	rs2.Fields("id_lang").Value		= CLng(Request.Form("cboLanguage").Item)
	rs2.Fields("datainscriere").Value = Now()
	If activeazacont Then rs2.Fields("datavalidare").Value = Now() 
	rs2.Update
	rs1.Update
	rs1.Close 
	rs2.Close
	Set rs1 = Nothing
	Set rs2 = Nothing
End Sub

Sub PrintFormResponse(strMessage, strMessageColor, bRetry)
%>
<html>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" style="background-color: ButtonFace" style="behavior:url('../_clientscripts/application.htc');">

<span unselectable="on" class=TLabel style="width:130px;height:13px;font-weight:bold;text-align:center;"
      style="left:100px;top:16px;">
Sign-up status
</span>
<span unselectable="on" class=TLabel style="width:323px;height:30px;font-weight:bold;"
      style="left:10px;top:40px;text-align:center;color:<%=strMessageColor%>;">
<%=strMessage%>
</span>
<%If not bRetry Then%>
<input id="btnNewRegRetry" type=button value="New sign-up" title="New sign-up"
       class=TButton style="width:100px;height:25px;"
       style="left:7px;top:80px;">
<%Else%>
<input id="btnNewRegRetry" type=button value="Retry" title="Retry to sign-up"
       class=TButton style="width:100px;height:25px;"
       style="left:7px;top:80px;">
<%End If%>
<input id="btnLogin" type=button value="Login" title="Login to application"
       class=TButton style="width:100px;height:25px;"
       style="left:117px;top:80px;">
<input id="btnClose" type=button value="Close" title="Close application"
       class=TButton style="width:100px;height:25px;"
       style="left:229px;top:80px;">

<script Language="VBScript">
Sub window_onLoad()
	If LCase(document.readyState) = "complete" Then 
		window.Parent.ShowConfirm
		btnNewRegRetry.Focus()
	End If 
End Sub

Sub btnNewRegRetry_onClick
	<%If not bRetry Then Response.Write "window.parent.ClearForm"%>	
	window.Parent.HideConfirm
End Sub

Sub btnLogin_OnClick
	window.parent.location = "login.asp"
End Sub

Sub btnClose_OnClick
	Call window.parent.parent.frames("Header").CloseApp
End Sub
</script>

</body>
</html>
<%
End Sub
%>
