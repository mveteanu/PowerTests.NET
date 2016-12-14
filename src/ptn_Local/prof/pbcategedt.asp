<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/problems.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim CategID, CursID
Dim CategDict, TotalDict
Dim ListBoxCateg, ListBoxCurs

CategID = Request.QueryString("CategID")
CursID  = Session("CursID")

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
Set CategDict = GetCategPBList(CategID, cn)
Set TotalDict = GetTotalPBList(CursID, cn)
ListBoxCateg = GetFillSelectFromDict(CategDict, "")
ListBoxCurs  = GetFillSelectFromDict(GetDictDifference(TotalDict,CategDict), "")
cn.Close
Set cn = nothing
%>
<html>
<head>
  <title>Edit category</title>
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
<fieldset class=TGroupBox style="width:485px;height:294px;"
          style="left:8px;top:8px;">
<legend>Edit category</legend>
<span class=TLabel style="width:105px;height:13px;"
      style="left:8px;top:24px;">
Included questions:
</span>
<span class=TLabel style="width:99px;height:13px;"
      style="left:272px;top:24px;">
Available questions:
</span>
<select id="ListBox1" size=2 multiple
        class=TListBox style="width:201px;height:230px;"
        style="left:8px;top:40px;">
<%=ListBoxCateg%>
</select>
<select id="ListBox2" size=2 multiple
        class=TListBox style="width:201px;height:230px;"
        style="left:272px;top:40px;">
<%=ListBoxCurs%>        
</select>
<a     class=TLinkLabel id=ButView1 href=# title="View selected questions"
       style="left:8px;top:270px;">View selected questions</a>
<a     class=TLinkLabel id=ButView2 href=# title="View selected questions"
       style="left:272px;top:270px;">View selected questions</a>

<input type=button value="<" title="Include selected questions"
       class=TButton style="width:25px;height:25px;"
       style="left:228px;top:80px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, "<"'>
<input type=button value="<<" title="Include all questions"
       class=TButton style="width:25px;height:25px;"
       style="left:228px;top:112px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, "<<"'>
<input type=button value=">" title="Remove selected questions"
       class=TButton style="width:25px;height:25px;"
       style="left:228px;top:160px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, ">"'>
<input type=button value=">>" title="Remove all questions"
       class=TButton style="width:25px;height:25px;"
       style="left:228px;top:192px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, ">>"'>

</fieldset>

<input id="Button1" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:168px;top:312px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:256px;top:312px;">
</div>


<script language=vbscript src="../_clientscripts/SelectControlUtils.vbs"></script>
<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Deschide o fereastra de preview cu problemele specificate prin 
' lista de ID-uri PBIDs
Sub PreviewProblems(PBIDs)
  If PBIDs<>"" then
   ShowModalDialog "pbviewscr.asp?PBIDs=" & PBIDs, "", "dialogWidth=720px;dialogHeight=500px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
  Else
   MsgBox "You need first to select one or more questions from above list.", vbOkOnly+vbExclamation
  End If 
End Sub

' Trateaza evenimentul care apare la apasare link-ului View... 1
Sub ButView1_onclick
  PreviewProblems GetItemsFromSelect(ListBox1, true)
End Sub

' Trateaza evenimentul care apare la apasare link-ului View... 2
Sub ButView2_onclick
  PreviewProblems GetItemsFromSelect(ListBox2, true)
End Sub

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  Dim itl
  
  itl = GetItemsFromSelect(ListBox1, false)
  If itl = "" then
    Window.returnValue = "<vid>"
  Else
    Window.returnValue = itl
  End If  
  Window.close 
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>
