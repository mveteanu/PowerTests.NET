<%@ Language=VBScript %>
<%
PrintPageHeader

if Request.QueryString.Count>0 then
  Session("CursID") = CLng(Request.QueryString("CursID"))
  PrintMenuLevel2
else
  PrintMenuLevel1
end if
%>


<%Sub PrintPageHeader%>
<html>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<%End Sub%>


<%Sub PrintMenuLevel1%>
<body style="background-color: buttonface" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0"
      unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<table class="window" cellspacing=0 cellpadding=0 WIDTH="100%" HEIGHT=54><tr>
<td valign=center align=left>
	<div id="butCursuri" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Gestionare cursuri">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile1.png"><br>Manage<br>courses&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div id="butSettings" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Setari generale">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/tools.png"><br>General<BR>settings&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
</td>
<td valign=center align=right>
	<div id="butExit" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Iesire din aplicatie">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/dooropen.png"><br>Exit</font>
		</td></tr></table>
	</div>
</td>
</tr></table>


<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language=vbscript>
Dim mymenu1

MenuItems1 = Array("Go to class","<HR>","Manage courses")
MenuItems2 = Array("Personal account settings")

sub butCursuri_onclick
 set mymenu1 = showmenu(0 , 52 ,150,"handlemenuclick1",MenuItems1)
end sub

sub butSettings_onclick
 set mymenu1 = showmenu(90 ,52,150,"handlemenuclick1",MenuItems2)
end sub

sub butExit_onclick
 CloseApp
end sub

sub handlemenuclick1(html)
 if html="<HR>" or html="" then exit sub
 mymenu1.Hide
 set mymenu1=nothing

 select case html
  case MenuItems1(0) window.parent.frames("Main").location = "entercours_cli.asp"
  case MenuItems1(2) window.parent.frames("Main").location = "cursurilist_cli.asp"
  
  case MenuItems2(0) window.parent.frames("Main").location = "../admin/modifyourself_cli.asp"
 end select
end sub

sub closecmd_onclick
   window.event.returnValue = false
   CloseApp
end sub

sub CloseApp
   If MsgBox("Are you sure you want to quit?",vbYesNo+vbQuestion,"Confirm") = vbYes Then window.parent.close
end sub
</script>

</body>
</html>
<%End Sub%>


<%Sub PrintMenuLevel2%>
<body style="background-color: buttonface" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">
<table class="window" cellspacing=0 cellpadding=0 WIDTH="100%" HEIGHT=54><tr>
<td valign=center align=left>
	<div id="butCursuri" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Setare curs curent">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile3.png"><br>Current<br>course&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div id="butStudenti" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Gestionare studenti">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/users.png"><br>Manage<br>students&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div id="butProbleme" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Gestionare probleme si categorii de probleme">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile1.png"><br>Manage<BR>questions&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div id="butTeste" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Gestionare teste si categorii de teste">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/crdfile1.png"><br>Manage<BR>tests&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
	&nbsp;
	<div id="butSettings" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Setari generale">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/tools.png"><br>General<BR>settings&nbsp;<font face="Webdings">6</font>
		</td></tr></table>
	</div>
</td>
<td valign=center align=right>
	<div id="butExit" class=TCoolButton onmousedown='this.style.border="inset thin"' onmouseup='this.style.border=""' onmouseover='this.style.border="outset thin"' onmouseout='this.style.border=""' title="Iesire din aplicatie">
		<table style="font-size:10px;" width=100% height=100% border=0 cellspacing=0 cellpadding=0><tr><td align=center valign=center>
		<img width=16 height=16 src="../images/dooropen.png"><br>Exit</font>
		</td></tr></table>
	</div>
</td>
</tr></table>


<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language=vbscript>
Dim mymenu1

MenuItems1 = Array("Course settings","<HR>","Close current course")
MenuItems2 = Array("Enrollment requests","<HR>","Enrolled students")
MenuItems3 = Array("Personal account settings")
MenuItems4 = Array("New question","Questions management","<HR>","Question categories management")
MenuItems5 = Array("Tests management","<HR>","Test categories management")

sub butCursuri_onclick
 set mymenu1 = showmenu(0 , 52 ,120,"handlemenuclick1",MenuItems1)
end sub

sub butStudenti_onclick
 set mymenu1 = showmenu(90 , 52 ,120,"handlemenuclick1",MenuItems2)
end sub

sub butProbleme_onclick
 set mymenu1 = showmenu(178 , 52 ,190,"handlemenuclick1",MenuItems4)
end sub

sub butTeste_onclick
 set mymenu1 = showmenu(267 , 52 ,180,"handlemenuclick1",MenuItems5)
end sub

sub butSettings_onclick
 set mymenu1 = showmenu(356 ,52,150,"handlemenuclick1",MenuItems3)
end sub

sub butExit_onclick
 CloseApp
end sub

sub handlemenuclick1(html)
 if html="<HR>" or html="" then exit sub
 mymenu1.Hide
 set mymenu1=nothing

 select case html     
  case MenuItems1(0) window.parent.frames("Main").location = "cursset_cli.asp"
  case MenuItems1(2) window.location.href = "headerprof.asp"
                     window.parent.frames("Main").location.href = "../login/middle.asp"
    
  case MenuItems2(0) window.parent.frames("Main").location = "cereristudlist_cli.asp"
  
  case MenuItems2(2) window.parent.frames("Main").location = "studlist_cli.asp"
  
  case MenuItems3(0) window.parent.frames("Main").location = "../admin/modifyourself_cli.asp"

  case MenuItems4(0) ShowModalDialog "pbcomposepb.asp", , "dialogWidth=768px;dialogHeight=547px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
                     
  case MenuItems4(1) window.parent.frames("Main").location = "pblist_cli.asp"
  
  case MenuItems4(3) window.parent.frames("Main").location = "pbcateglist_cli.asp"

  case MenuItems5(0) window.parent.frames("Main").location = "tstlist_cli.asp"
  
  case MenuItems5(2) window.parent.frames("Main").location = "tstcateglist_cli.asp"
 end select
end sub

sub document_oncontextmenu
   window.event.returnValue = false
end sub

sub closecmd_onclick
   window.event.returnValue = false
   CloseApp
end sub

sub CloseApp
   If MsgBox("Are you sure you want to quit?",vbYesNo+vbQuestion,"Confirm") = vbYes Then window.parent.close
end sub
</script>

</body>
</html>
<%End Sub%>
