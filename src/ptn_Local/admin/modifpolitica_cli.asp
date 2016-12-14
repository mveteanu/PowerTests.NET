<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/settings.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Dim UserPolicy(3)
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 UserPolicy(0) = GetUserPolicy("A",cn)
 UserPolicy(1) = GetUserPolicy("P",cn)
 UserPolicy(2) = GetUserPolicy("S",cn)
 cn.Close
 Set cn = nothing
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:370px;height:250px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:370px;height:250px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<form name="FormularS">
<span id="Label1"
      class=TLabel style="width:223px;height:13px; font-weight:bold;"
      style="left:94px;top:16px;">
New user acceptance policy
</span>
<span id="Label2"
      class=TLabel style="width:65px;height:13px;"
      style="left:56px;top:72px;">
Administrators:
</span>
<span id="Label3"
      class=TLabel style="width:44px;height:13px;"
      style="left:56px;top:112px;">
Professors:
</span>
<span id="Label4"
      class=TLabel style="width:42px;height:13px;"
      style="left:56px;top:152px;">
Students:
</span>
<select id="ComboBox1"
        class=TComboBox style="width:190px;"
        style="left:133px;top:64px;">
<option value=0 <%=CompareToString(0,UserPolicy(0)," SELECTED ","")%> >Accept new requests automatically</option>
<option value=1 <%=CompareToString(1,UserPolicy(0)," SELECTED ","")%> >Require administrator acceptance</option>
<option value=2 <%=CompareToString(2,UserPolicy(0)," SELECTED ","")%> >Deny all new requests</option>
</select>
<select id="ComboBox2"
        class=TComboBox style="width:190px;"
        style="left:133px;top:104px;">
<option value=0 <%=CompareToString(0,UserPolicy(1)," SELECTED ","")%> >Accept new requests automatically</option>
<option value=1 <%=CompareToString(1,UserPolicy(1)," SELECTED ","")%> >Require administrator acceptance</option>
<option value=2 <%=CompareToString(2,UserPolicy(1)," SELECTED ","")%> >Deny all new requests</option>
</select>
<select id="ComboBox3"
        class=TComboBox style="width:190px;"
        style="left:133px;top:144px;">
<option value=0 <%=CompareToString(0,UserPolicy(2)," SELECTED ","")%> >Accept new requests automatically</option>
<option value=1 <%=CompareToString(1,UserPolicy(2)," SELECTED ","")%> >Require administrator acceptance</option>
<option value=2 <%=CompareToString(2,UserPolicy(2)," SELECTED ","")%> >Deny all new requests</option>
</select>

<input id="Button1" type=button value="Save" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:92px;top:200px;">

<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:204px;top:200px;">
</form>
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="modifpolitica_ser.asp" target="FormReturn">
<input id="ComboBox1" name="adminpolicy" type=text>
<input id="ComboBox2" name="profpolicy" type=text>
<input id="ComboBox3" name="studpolicy" type=text>
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language="vbscript" src="../_clientscripts/formutils.vbs"></script>
<script language=vbscript>
' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub window_onload
  if not CopyForms(FormularS,FormularH) then _
    msgbox "Error loading user data!"&vbCrLf&"Please contact a system administrator.", vbOkOnly+vbCritical
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub



' Schimba starea activ/inactiv a butoanelor de Save si Close
' Daca state = true butoanele sunt active si invers
Sub ActivateButtons(state)
  FormularS.Button1.disabled = not state
  FormularS.Button2.disabled = not state
End Sub


' Trateaza evenimentul ce apare la apasarea butonului Save
Sub Button1_OnClick
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
   if msgbox("Form data was changed."&vbcrlf&vbcrlf&"Do you want to save changes before closing the form?", vbYesNo+vbQuestion) = vbYes then
     Button1_OnClick
   else
     HideAllDivs
   end if     
 end if
End Sub
</script>

</BODY>
</HTML>
