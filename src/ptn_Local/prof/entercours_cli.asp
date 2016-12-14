<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/cursuri.asp" -->
<%
Sub FillCursuriListBox
 set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 set rs = GetCursuriByProf(Session("UserID"),cn)
 do until rs.EOF
  Response.Write "<option value="& rs.Fields("id_curs").Value &">"& rs.Fields("numecurs").Value &"</option>"& vbCrLf
  rs.MoveNext
 loop
 rs.Close
 set rs=nothing
 cn.Close 
 set cn=nothing
End Sub
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:406px;height:263px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:406px;height:263px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<span id="Label1"
      class=TLabel style="width:75px;height:13px;"
      style="left:8px;top:8px;">
Select course:
</span>
<select id="ListBox1" 
        size=2
        class=TListBox style="width:385px;height:195px;"
        style="left:8px;top:24px;">
<%FillCursuriListBox%>
</select>
<input id="Button1" type=button value="Manage courses" title="Manage courses"
       class=TButton style="width:97px;height:25px;"
       style="left:9px;top:224px;">
<input id="Button2" type=button value="Cancel" title="Cancel"
       class=TButton style="width:75px;height:25px;"
       style="left:320px;top:224px;">
<input id="Button3" type=button value="Go to class" title="Go to selected course"
       class=TButton style="width:75px;height:25px;"
       style="left:240px;top:224px;">
</div>

<script language=vbscript>
' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub


' Trateaza evenimentul care apare la apasarea butonului Gestionare cursuri
Sub Button1_onclick
  Window.location.href = "cursurilist_cli.asp"
End Sub


' Trateaza evenimentul care apare la apasarea butonului Cancel
Sub Button2_onclick
  HideAllDivs
End Sub

' Trateaza evenimentul care apare la apasarea butonului Intra la curs
Sub Button3_onclick
  if ListBox1.Value="" then
    msgbox "You need to select a course from the list.",vbOkOnly+vbExclamation, "PowerTests .NET"
  else
    window.parent.frames("Header").location.href = "headerprof.asp?CursID=" & ListBox1.Value
    HideAllDivs
  end if
End Sub
</script>

</body>
</html>

