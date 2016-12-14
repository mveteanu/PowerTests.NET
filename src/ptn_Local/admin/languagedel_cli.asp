<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/settings.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	
	Dim langid
	
	langid = CLng(Request.QueryString("id").Item)
	nruser = CLng(Request.QueryString("nru").Item)

	Function LanguageOptions
		Dim re
		Set db = new clsDatabase
		re = GetOptionsForSelect("<option value='@0' @SELECTED>@1</option>", "@0 = " & GetPreferredLanguage(db.Connection), "SELECT * FROM TBLanguages WHERE id<>" & langid, Array("id", "langname"), db.Connection)
		Set db = Nothing
		LanguageOptions = re
	End Function
%>
<html>
<head>
  <title>Delete language</title>
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
<span class=TLabel style="width:240px;height:40px;"
      style="left:8px;top:12px;">
Selected language is used by <%=nruser%> application users. Select the new language that will be automatically assigned to them after deletion.
</span>
<span class=TLabel style="width:80px;height:13px;"
      style="left:16px;top:68px;">
New language:
</span>
<select id="cboLanguage" name="cboLanguage"
		class=TComboBox
		style="left:106px;top:66px;"
		style="width:145px;height:21px;">
		<%=LanguageOptions()%>
</select>

<input id="Button1" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:57px;top:104px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:161px;top:104px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
	Form1.style.visibility = "visible"
	WaitforForm.style.visibility = "hidden"
End Sub

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
	If cboLanguage.selectedIndex = -1 then 
		MsgBox "New language was not selected.", vbOkOnly+vbExclamation, "Warning"
	Else
		window.returnValue = cboLanguage(cboLanguage.selectedIndex).value
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
