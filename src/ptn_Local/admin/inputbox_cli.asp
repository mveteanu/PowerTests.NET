<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/languages.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	
	strPagetitle	= Request.QueryString("pagetitle").Item 
	strLabel		= Request.QueryString("label").Item 
	strTextValue	= Request.QueryString("text").Item 
	bAcceptEmpty	= CBool(Len(Request.QueryString("acceptempty").Item) > 0)
%>
<html>
<head>
  <title><%=strPagetitle%></title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body>

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
<span class=TLabel style="width:250px;height:13px;"
      style="left:8px;top:16px;">
<%=strLabel%>
</span>
<input id="Edit1" type=text maxlength=50 value="<%=strTextValue%>"
       class=TEdit style="width:257px;height:21px;"
       style="left:8px;top:32px;">
<input id="Button1" type=button value="OK" title="Accept changes"
       class=TButton style="width:75px;height:25px;"
       style="left:55px;top:80px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:143px;top:80px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
	Form1.style.visibility = "visible"
	WaitforForm.style.visibility = "hidden"
	Edit1.focus 
End Sub

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
	If bAcceptEmpty Then
		window.returnValue = Edit1.value
		window.close
	Else
		If Len(Trim(Edit1.value)) = 0 Then 
			MsgBox "Invalid text.", vbOkOnly+vbExclamation, "Warning"
		Else
			window.returnValue = Edit1.value
			window.close
		End If
	End If
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
	Window.close 
End Sub
</script>

</body>
</html>
