<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/problems.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim categid, comparid, problnrVal
Dim categLst, comparLst

If Request.QueryString.Count = 0 then
 categid    = ""
 comparid   = ""
 problnrVal = ""
Else
 categid    = Request.QueryString("p1")
 comparid   = Request.QueryString("p2")
 problnrVal = Request.QueryString("p3")
End If

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
categLst  = GetFillSelectFromDict(GetCategDiction(Session("CursID"), cn), categid)
comparLst = GetFillSelectFromDict(DictFromTwoArrays(Array(0,1,2),Array("Fix","Min","Max")), comparid)
cn.Close
Set cn=nothing
%>
<html>
<head>
  <title>Specify criteria</title>
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
<fieldset class=TGroupBox style="width:313px;height:129px;"
          style="left:8px;top:8px;">
<legend>Define filter</legend>
<span class=TLabel style="width:48px;height:13px;"
      style="left:16px;top:32px;">
Category:
</span>
<span class=TLabel style="width:70px;height:13px;"
      style="left:16px;top:64px;">
Comparation:
</span>
<span class=TLabel style="width:93px;height:13px;"
      style="left:16px;top:96px;">
No questions:
</span>
<select id="ComboBox1"
        class=TComboBox style="width:193px;"
        style="left:104px;top:24px;">
<%=categLst%>
</select>
<select id="ComboBox2"
        class=TComboBox style="width:193px;"
        style="left:104px;top:56px;">
<%=comparLst%>
</select>
<input id="Edit1" type=text maxlength=4
       class=TEdit style="width:193px;height:21px;"
       style="left:104px;top:88px;"
       value="<%=problnrVal%>">
</fieldset>
<input id="Button1" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:83px;top:152px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:171px;top:152px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului OK
Sub Button1_onclick
  Dim si1, si2
  
  si1 = ComboBox1.selectedIndex 
  si2 = ComboBox2.selectedIndex 
  If (not IsNumeric(Edit1.Value)) or (si1=-1) or (si2=-2) then
    msgbox "Nu ati completat corect datele solicitate", vbOkOnly+vbExclamation
  Else
    Window.returnValue = ComboBox1.options(si1).Value & "|" &_
                         ComboBox2.options(si2).Value & "|" &_ 
                         Edit1.Value & "|" &_ 
                         ComboBox1.options(si1).Text
    Window.close 
  End If  
End Sub


' Evenimentul apare la apasarea butonului Close
Sub Button2_onclick
  Window.returnValue = ""
  Window.close 
End Sub
</script>

</body>
</html>
