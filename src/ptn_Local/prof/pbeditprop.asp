<%@ Language=VBScript %>
<html>
<head>
  <title>Answer properties</title>
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


<input id="RadioButton1" name=radiomaingrp type=radio CHECKED
       class=TButton
       style="left:8px;top:16px;"
       onclick="vbscript:ToggleGroupbox">
<input id="RadioButton2" name=radiomaingrp type=radio
       class=TButton
       style="left:8px;top:48px;"
       onclick="vbscript:ToggleGroupbox">
<label for=RadioButton1 
       class=TLabel style="width:80px;height:13px;"
       style="cursor:hand; left:32px;top:20px;">Numeric
</label>
<label for=RadioButton2
       class=TLabel style="width:80px;height:13px;"
       style="cursor:hand; left:32px;top:52px;">Text
</label>


<fieldset id=GroupBox1 class=TGroupBox 
          style="width:249px;height:161px;"
          style="left:120px;top:8px;">
<legend>Numeric answer properties</legend>
<span class=TLabel style="width:43px;height:13px;"
      style="left:16px;top:32px;">
Precision:
</span>
<select id="ComboBox1"
        class=TComboBox style="width:155px;"
        style="left:72px;top:32px;">
<option value=0>0 decimals</option>
<option value=1>1 decimal</option>
<option value=2 SELECTED>2 decimals</option>
<option value=3>3 decimals</option>
<option value=4>4 decimals</option>
<option value=5>5 decimals</option>
<option value=6>6 decimals</option>
<option value=7>All decimals</option>
</select>
<input id="Checkbox1" type=checkbox
       class=TButton
       style="left:16px;top:80px;">
<label for=Checkbox1
       class=TLabel style="width:175px;height:13px;"
       style="cursor:hand; left:40px;top:84px;">Ignore non-numeric characters
</label>
</fieldset>


<fieldset id=GroupBox2 class=TGroupBox 
          style="visibility:hidden;"
          style="width:249px;height:161px;"
          style="left:120px;top:8px;">
<legend>Text answer properties</legend>
<span class=TLabel style="width:41px;height:13px;"
      style="left:16px;top:32px;">
Match:
</span>
<select id="ComboBox2"
        class=TComboBox style="width:155px;"
        style="left:72px;top:32px;">
<option value=0 SELECTED>Exact match</option>
<option value=1>Begins with</option>
<option value=2>Ends with</option>
<option value=3>Contains</option>
</select>
<input id="Checkbox2" type=checkbox
       class=TButton
       style="left:16px;top:80px;">
<input id="Checkbox3" type=checkbox
       class=TButton
       style="left:16px;top:120px;">
<label for=Checkbox2
       class=TLabel style="width:190px;height:13px;"
       style="cursor:hand; left:40px;top:84px;">Case sensitive
</label>
<label for=Checkbox3
       class=TLabel style="width:175px;height:13px;"
       style="cursor:hand; left:40px;top:124px;">Ignore trailing spaces
</label>
</fieldset>


<input id="Button1" type=button value="OK" title="OK"
       class=TButton style="width:75px;height:25px;"
       style="left:212px;top:184px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:294px;top:184px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  DecodeInputValues
End Sub


' Decodeaza informatiile trimise prin DialogArguments si 
' le foloseste pentru setarea elementelor din formular
Sub DecodeInputValues
  Dim i
  
  i = CInt(Window.DialogArguments)
  If ((i and 1) = 0) then   ' Raspuns numeric
    RadioButton1.checked = true
    ComboBox1.selectedIndex = (i and 14)/2
    Checkbox1.checked = ((i and 16) = 16)
  else                  ' Raspuns alfanumeric
    RadioButton2.checked = true
    ComboBox2.selectedIndex = (i and 14)/2
    Checkbox2.checked = ((i and 16) = 16)
    Checkbox3.checked = ((i and 32) = 32) 
  end if
  call ToggleGroupbox
End Sub


' Obtine starea controalelor din fereastra in vederea inpachetarii
' informatiilor in parametrul ReturnValue la apasarea butonului OK
Function GetOutputValues
 Dim i
 
 If RadioButton1.checked then 
   i = ComboBox1.selectedIndex * 2
   If Checkbox1.checked then i = i + 16
 else 
   i = 1 + ComboBox2.selectedIndex * 2
   If Checkbox2.checked then i = i + 16
   If Checkbox3.checked then i = i + 32
 end if
 
 GetOutputValues = i
End Function


' Comuta afisarea celor 2 groupbox-uri la actionarea radiobutoanelor
Sub ToggleGroupbox
  If RadioButton1.checked then
    GroupBox1.style.visibility = ""
    GroupBox2.style.visibility = "hidden"
  Else
    GroupBox1.style.visibility = "hidden"
    GroupBox2.style.visibility = ""
  End If    
End Sub



' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  Window.returnValue = GetOutputValues
  Window.close 
End Sub


' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>