<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/TabControl.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<%
	Dim PersonData(5), lang
 
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open Application("DSN")
	Set RSUser = GetUserByID(Session("UserID"),cn)
	PersonData(0) = RSUser.Fields("nume").Value
	PersonData(1) = RSUser.Fields("prenume").Value
	PersonData(2) = RSUser.Fields("email").Value
	PersonData(3) = RSUser.Fields("telefon").Value
	lang		  = RSUser.Fields("id_lang").Value
	
	PersonData(4) = GetOptionsForSelect("<option value='@0' @SELECTED>@1</option>", "@0 = " & lang, "TBLanguages", Array("id", "langname"), cn)
	
	RSUser.Close
	Set RSUser=nothing
	cn.Close 
	Set cn=nothing
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
 <script language=javascript src="../_clientscripts/tabControlEvents.js"></script>
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:410px;height:300px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:410px;height:300px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<form name="FormularS">
<%OpenTabControl 11,8,380, 241, Array("Personal info","Password"), 1, "PageControl1"%>
<%OpenTabContent%>
<span id="Label1"
      class=TLabel style="width:110px;height:13px;"
      style="left:60px;top:40px;">
Last name:
</span>
<span id="Label2"
      class=TLabel style="width:110px;height:13px;"
      style="left:60px;top:72px;">
First name:
</span>
<span id="Label3"
      class=TLabel style="width:110px;height:13px;"
      style="left:60px;top:104px;">
Email:
</span>
<span id="Label4"
      class=TLabel style="width:110px;height:13px;"
      style="left:60px;top:136px;">
Phone:
</span>
<span id="Label5"
      class=TLabel style="width:110px;height:13px;"
      style="left:60px;top:168px;">
Preferred language:
</span>
<input id="Edit1" type=text maxlength=20 value="<%=PersonData(0)%>"
       class=TEdit style="width:121px;height:21px;"
       style="left:174px;top:32px;">
<input id="Edit2" type=text maxlength=20 value="<%=PersonData(1)%>"
       class=TEdit style="width:121px;height:21px;"
       style="left:174px;top:64px;">
<input id="Edit3" type=text maxlength=50 value="<%=PersonData(2)%>"
       class=TEdit style="width:121px;height:21px;"
       style="left:174px;top:96px;">
<input id="Edit4" type=text maxlength=20 value="<%=PersonData(3)%>"
       class=TEdit style="width:121px;height:21px;"
       style="left:174px;top:128px;">
<select id="Edit8"
		class=TComboBox
		style="left:174px;top:160px;"
		style="width:121px;height:21px;">
		<%=PersonData(4)%>
</select>
<%CloseTabContent%>
<%OpenTabContent%>
<span id="Label5"
      class=TLabel style="width:120px;height:13px;"
      style="left:48px;top:48px;">
Old password:
</span>
<span id="Label6"
      class=TLabel style="width:120px;height:13px;"
      style="left:48px;top:80px;">
New password:
</span>
<span id="Label7"
      class=TLabel style="width:120px;height:13px;"
      style="left:48px;top:112px;">
Verify new password:
</span>
<input id="Edit5" type=password maxlength=20
       class=TEdit style="width:121px;height:21px;"
       style="left:184px;top:40px;">
<input id="Edit6" type=password maxlength=20
       class=TEdit style="width:121px;height:21px;"
       style="left:184px;top:72px;">
<input id="Edit7" type=password maxlength=20
       class=TEdit style="width:121px;height:21px;"
       style="left:184px;top:104px;">
<%CloseTabContent%>
<%CloseTabControl%>

<input id="Button1" type=button value="Save" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:116px;top:260px;">

<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:212px;top:260px;">
</form>
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="modifyourself_ser.asp" target="FormReturn">
<input id="Edit1" name="nume" type=text>
<input id="Edit2" name="prenume" type=text>
<input id="Edit3" name="email" type=text>
<input id="Edit4" name="telefon" type=text>
<input id="Edit5" name="oldpass" type=text>
<input id="Edit6" name="newpass" type=text>
<input id="Edit7" name="newpassver" type=text>
<input id="Edit8" name="EditLang" type=text>
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language="javascript" src="../_clientscripts/emails.js"></script>
<script language="vbscript" src="../_clientscripts/formutils.vbs"></script>
<script language=vbscript>
' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub window_onload
  if not CopyForms(FormularS,FormularH) then _
    msgbox "Error loading user data!"&vbCrLf&"Contact a system administrator.", vbOkOnly+vbCritical
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub



' Schimba starea activ/inactiv a butoanelor de Save si Close
' Daca state = true butoanele sunt active si invers
Sub ActivateButtons(state)
  FormularS.Button1.disabled = not state
  FormularS.Button2.disabled = not state
End Sub

' Verifica corectitudinea informatiilor din campurile pentru
' schimbarea parolei
Function VerifyPasswordTab
 Dim re
 If FormularS.Edit6.value <> FormularS.Edit7.value then
   re = 1 ' Parola noua <> Verificare parola noua
 ElseIf FormularS.Edit6.value <> "" and  FormularS.Edit5.value = "" then
   re = 2 ' S-a introdus noua parola dar nu s-a introdus vechea parola
 ElseIf FormularS.Edit5.value <> "" and FormularS.Edit6.value = "" then
   re = 3 ' S-a introdus vechea parola dar nu s-a specificat noua parola  
 Else
   re = 0 ' Totul pare OK
 End If
 VerifyPasswordTab = re  
End Function

' Verifica corectitudinea informatiilor din campurile pentru
' schimbarea datelor utilizatorului
Function VerifyUserData
 Dim re
 re = true
 if FormularS.Edit1.value = "" then
  re = 1 ' Nume null
 elseif FormularS.Edit2.value = "" then
  re = 2 ' Prenume null
 elseif not validEMail(FormularS.Edit3.value) then
  re = 3 ' Adresa de email incorecta
 end if 
 VerifyUserData = re
End Function

' Trateaza evenimentul ce apare la apasarea butonului Save
Sub Button1_OnClick
 Select Case VerifyPasswordTab
   case 1 MsgBox "Password verification is different than password."&vbCrLf&"Please type carefully.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrParola_cont,tabCntParola_cont
          Exit Sub
   case 2 MsgBox "In order to change the password you need to specify the old one as well."&vbCrLf&"This is a safety measure.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrParola_cont,tabCntParola_cont
          Exit Sub
   case 3 MsgBox "New password cannot be empty."&vbCrLf&"Make sure you type a strong password.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrParola_cont,tabCntParola_cont
          Exit Sub
 End Select

 Select Case VerifyUserData
   case 1 MsgBox "Last name cannot be empty."&vbCrLf&"Cannot save data.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrDate_user,tabCntDate_user
          Exit Sub
   case 2 MsgBox "First name cannot be empty."&vbCrLf&"Cannot save data.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrDate_user,tabCntDate_user
          Exit Sub
   case 3 MsgBox "Invalid email address."&vbCrLf&"Cannot save data.", vbOkOnly+vbExclamation
          tabActivate tabPageControl1,tabHdrDate_user,tabCntDate_user
          Exit Sub
 End Select

 if not CopyForms(FormularS,FormularH) then
    msgbox "Error saving data!"&vbCrLf&"Please contact a system administrator.", vbOkOnly+vbCritical
    Exit Sub
 end if

 ActivateButtons false
 FormularH.Submit
End Sub


' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close. O alta strategie ar fi constat
' in navigarea catre o pagina goala prin apelarea unei metodode
' de tipul: call Window.parent.frames("Header").NavigateToMain
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Trateaza evenimentul ce apare la apasarea butonului Close
Sub Button2_OnClick
 if CompareForms(FormularS,FormularH) then 
   HideAllDivs
 else
   if msgbox("Data was changed."&vbcrlf&vbcrlf&"Do you want to save changes before closing the form?", vbYesNo+vbQuestion) = vbYes then
     Button1_OnClick
   else
     HideAllDivs
   end if     
 end if
End Sub
</script>

</BODY>
</HTML>
