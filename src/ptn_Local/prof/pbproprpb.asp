<%@ Language=VBScript %>
<%
Response.Buffer = True
Response.Expires = -1

If Request.QueryString.Count = 0 then
  PrintEditProperties
Else
  set cn = Server.CreateObject("ADODB.Connection") 
  cn.Open Application("DSN")
  PrintRenameProblem(GetProblemName(Request.QueryString("PBId"), cn))
  cn.Close 
  set cn = nothing
End If

' Intoarce numele problemei in fn. de ID-ul ei
Function GetProblemName(pbid, objCon)
 Dim re
 Dim myCmd, Rs
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetPbInfoByID"
 myCmd.CommandType = adCmdStoredProc
 Set Rs = myCmd.Execute(,CLng(pbid))
 If not Rs.EOF then
   re = Rs.Fields("numeproblema").Value
 Else
   re = ""  
 End If
 Rs.Close
 set rs = nothing
 set myCmd = nothing
 
 GetProblemName = re
End Function


' Creaza fereastra pentru editarea proprietatilor problemei
' Fereastra e afisata de editorul de probleme
Sub PrintEditProperties
%>
<html>
<head>
  <title>Question properties</title>
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
<fieldset id=GroupBox1 class=TGroupBox 
          style="width:305px;height:153px;"
          style="left:8px;top:8px;">
<legend>Question properties</legend>
<span class=TLabel style="width:31px;height:13px;"
      style="left:16px;top:40px;">
Name:
</span>
<span class=TLabel style="width:36px;height:13px;"
      style="left:16px;top:80px;">
Author:
</span>
<input id="Edit1" type=text maxlength=50
       class=TEdit style="width:200px;height:21px;"
       style="left:64px;top:32px;">
<input id="Edit2" type=text maxlength=50
       class=TEdit style="width:200px;height:21px;"
       style="left:64px;top:72px;">
<input id="Checkbox1" type=checkbox
       class=TButton
       style="left:16px;top:120px;">
<label for=Checkbox1
       class=TLabel style="width:150px;height:13px;"
       style="cursor:hand; left:40px;top:124px;">Allow partial answer
</label>
</fieldset>

<input id="Button1" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:79px;top:176px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:167px;top:176px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  ApplyDialogArguments
End Sub

' Aplica controalelor din fereastra valorile trimise prin
' DialogArguments
Sub ApplyDialogArguments
 Dim da
 
 da = Split(window.dialogArguments, Chr(3), -1, 1)
 Edit1.value = da(0)
 Edit2.value = da(1)
 select case LCase(da(2))
   case "d" Checkbox1.disabled = true
   case "t" Checkbox1.checked = true
   case else Checkbox1.checked = false
 end select
End Sub

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  Dim r1, r2, r3
  r1 = Edit1.value
  r2 = Edit2.value
  if Checkbox1.disabled then 
    r3 = "d"
  elseif Checkbox1.checked then
    r3 = "t"
  else
    r3 = "f"
  end if    
  window.returnValue = r1 & Chr(3) & r2 & Chr(3) & r3
  Window.close 
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>
<%
End Sub ' PrintEditProperties


' Creaza fereastra pentru redenumirea problemei
' Fereastra e afisata de Problem Manager
Sub PrintRenameProblem(numevechi)
%>
<html>
<head>
  <title>Rename question</title>
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
  If Trim(Edit1.value) = "" then 
    MsgBox "You need to enter question name.", vbOkOnly+vbExclamation, "Warning"
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
<%
End Sub ' PrintRenameProblem
%>
