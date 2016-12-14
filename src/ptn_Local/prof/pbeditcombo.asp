<%@ Language=VBScript %>
<html>
<head>
  <title>List items</title>
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

<textarea id="Memo1"
       class=TEdit style="width:180px;height:137px;"
       style="overflow:auto;"
       style="left:6px;top:8px;">
</textarea>

<input id="Button1" type=button value="OK" title="OK changes"
       class=TButton style="width:50px;height:17px;"
       style="left:39px;top:148px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:50px;height:17px;"
       style="left:103px;top:148px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  If IsArray(Window.DialogArguments) then
    Memo1.innerText = Join(Window.DialogArguments,vbCrLf)
  End If   
  Memo1.focus
End Sub


' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  Window.returnValue = Split(Memo1.innerText, vbCrLf, -1, 1)
  Window.close 
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>