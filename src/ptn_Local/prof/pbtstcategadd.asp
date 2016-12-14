<%@ Language=VBScript %>
<%
Response.Buffer = True
Response.Expires = -1

Dim cn, querycategid
Dim querytip, oldcategname

querycategid = Request.QueryString("CategID")
If querycategid = "" then
  PrintWindow "Add category", ""
Else
  querytip = Request.QueryString("TipCateg")
  set cn = Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  If LCase(querytip) = "tst" then
    oldcategname = GetCategName("tst", querycategid, cn)
  else
    oldcategname = GetCategName("pb", querycategid, cn)
  end if
  PrintWindow "Rename category", oldcategname
  cn.Close
  set cn = nothing
End If


' Obtine numele unei categorii de probleme in functie de ID-ul acesteia
Function GetCategName(TipCateg, CategID, objCon)
 Dim rs, re, SQLSel
 Const SQLSelPb  = "SELECT * FROM TBProblemsCategories WHERE id_categpb=@1"
 Const SQLSelTst = "SELECT * FROM TBTestsCategories WHERE id_categtst=@1"
 If LCase(TipCateg) = "tst" then SQLSel = SQLSelTst else SQLSel = SQLSelPb

 Set rs = objCon.Execute(Replace(SQLSel, "@1", CStr(CategID)))
 If not rs.EOF then
   re = rs.Fields("numecateg").Value
 Else
   re = ""
 End If
 rs.Close
 set rs = nothing
 GetCategName = re
End Function


Sub PrintWindow(titlu, numevechi)
%>
<html>
<head>
  <title><%=titlu%></title>
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
<input id="Edit1" type=text maxlength=49 value="<%=numevechi%>"
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
    MsgBox "You need to enter category name.", vbOkOnly+vbExclamation, "Warning"
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
End Sub ' PrintWindow
%>
