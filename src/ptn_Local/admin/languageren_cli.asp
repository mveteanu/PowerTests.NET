<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/languages.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	
	Dim langid
	Dim numevechi
	
	langid	  = CLng(Request.QueryString("id").Item)
	numevechi = GetLanguageByID(langid, Nothing)
%>
<html>
<head>
  <title>Rename language</title>
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
<span class=TLabel style="width:31px;height:13px;"
      style="left:8px;top:16px;">
Name:
</span>
<input id="Edit1" type=text maxlength=50 value="<%=numevechi%>"
       class=TEdit style="width:257px;height:21px;"
       style="left:8px;top:32px;">
<input id="Button1" type=button value="OK" title="Save changes"
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
	If Len(Trim(Edit1.value)) = 0 then 
		MsgBox "Language name cannot be empty.", vbOkOnly+vbExclamation, "Warning"
	Else
		window.returnValue = Edit1.value
		window.close 
	End If  
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
	Window.close 
End Sub
</script>

</body>
</html>
