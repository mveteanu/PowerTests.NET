<%@ Language=VBScript %>
<html>
<head>
  <title>Answer type</title>
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

<span id="Label1"
      class=TLabel style="width:89px;height:39px;"
      style="left:16px;top:24px;">
Specify desired<br>answer type<br>then press OK
</span>
<fieldset class=TGroupBox style="width:226px;height:161px;"
          style="left:128px;top:16px;">
<legend>Answer type</legend>
<input id="RadioButton1" name=radiogrp1 type=radio value="1" CHECKED
       class=TButton
       style="left:10px;top:30px;">
<input id="RadioButton2" name=radiogrp1 type=radio value="2"
       class=TButton
       style="left:10px;top:60px;">
<input id="RadioButton3" name=radiogrp1 type=radio value="3"
       class=TButton
       style="left:10px;top:90px;">
<input id="RadioButton4" name=radiogrp1 type=radio value="4"
       class=TButton
       style="left:10px;top:120px;">
<label for=RadioButton1 
       class=TLabel style="width:180px;height:13px;"
       style="cursor:hand; left:32px;top:34px;">Simple selection - Radio
</label>
<label for=RadioButton2
       class=TLabel style="width:180px;height:13px;"
       style="cursor:hand; left:32px;top:64px;">Multiple selection - Check
</label>
<label for=RadioButton3
       class=TLabel style="width:180px;height:13px;"
       style="cursor:hand; left:32px;top:94px;">Multi-combined selection - Combo
</label>
<label for=RadioButton4 
       class=TLabel style="width:180px;height:13px;"
       style="cursor:hand; left:32px;top:124px;">Data entry - Edit
</label>
</fieldset>
<input id="Button1" type=button value="OK" title="Add answer"
       class=TButton style="width:75px;height:25px;"
       style="left:199px;top:190px;">
<input id="Button2" type=button value="Cancel" title="Cancel add answer"
       class=TButton style="width:75px;height:25px;"
       style="left:279px;top:190px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub


Function GetCheckedControl
  Dim re
  If RadioButton1.checked then
    re = 1
  elseif RadioButton2.checked then
    re = 2
  elseif RadioButton3.checked then
    re = 3
  elseif RadioButton4.checked then
    re = 4
  else
    re = 0
  end if   
  GetCheckedControl = re
End Function

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  Window.returnValue = GetCheckedControl
  Window.close 
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>