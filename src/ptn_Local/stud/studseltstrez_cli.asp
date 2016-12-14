<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/TabControl.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

UserID           = Request.QueryString("userid")
ShowInBrowserWin = false
If UserID = "" then 
  UserID = Session("UserID")
  ShowInBrowserWin = true
End If  
%>
<HTML>
<head>
 <title>Select solved tests</title>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language="vbscript" src="../_clientscripts/utils.vbs"></script>
<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<script language=vbscript src="../_clientscripts/MiscControlUtils.vbs"></script>
<script language=javascript src="../_clientscripts/tabControlEvents.js"></script> 

<%If ShowInBrowserWin then%>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:500px;height:360px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%Else%>
<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<%End If%>
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<%If ShowInBrowserWin then%>
<div id="Form1" style="visibility:hidden;" unselectable="on"
     class=TForm style="width:500px;height:360px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%Else%>
<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<%
 End If
 
 OpenTabControl 8,8,473, 305, Array("Tests","Options"), 1, "PageControl1"
 OpenTabContent
 CreateTableControl 8, 8, 259, Array("Test name","Completed","Abandoned"), Array(229,110,110), 2, "studseltstrez_dat.asp?userid=" & CStr(UserID) , true, "MyTable" 
 CloseTabContent
 OpenTabContent
%>
<fieldset class=TGroupBox style="width:441px;height:250px;cursor:default;"
          style="left:8px;top:8px;">
<legend>Test solving date</legend>

<input id="RadioButton1" name=radiomaingrp type=radio CHECKED
       class=TButton style="left:8px;top:20px;"
       onclick="vbscript:CheckDateRadios">
<label for=RadioButton1 unselectable='on'
       class=TLabel style="width:177px;height:13px;"
       style="cursor:hand; left:32px;top:24px;">All dates
</label>

<input id="RadioButton2" name=radiomaingrp type=radio
       class=TButton style="left:8px;top:54px;"
       onclick="vbscript:CheckDateRadios">
<label for=RadioButton2  unselectable='on'
       class=TLabel style="width:201px;height:13px;"
       style="cursor:hand; left:32px;top:58px;">After date
</label>

<input id="RadioButton3" name=radiomaingrp type=radio
       class=TButton style="left:8px;top:136px;"
       onclick="vbscript:CheckDateRadios">
<label for=RadioButton3  unselectable='on'
       class=TLabel style="width:185px;height:13px;"
       style="cursor:hand; left:32px;top:140px;">Between dates
</label>
<%
CreateSimpleDateSelector 2000, 2020, 32, 70,  "data1"
CreateSimpleDateSelector 2000, 2020, 32, 150, "data2"
CreateSimpleDateSelector 2000, 2020, 32, 190, "data3"
%>
</fieldset>

<%CloseTabContent%>
<%CloseTabControl%>

<input DISABLED id="Button1" type=button value="View results" title="View results for selected tests"
       class=TButton style="width:100px;height:25px;"
       style="left:132px;top:322px;">
<input id="Button2" type=button value="Close" title="Close form"
       class=TButton style="width:100px;height:25px;"
       style="left:260px;top:322px;">

</div>

<script language=vbscript>
' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"

   SimpleControlDisabled "data1", true
   SimpleControlDisabled "data2", true
   SimpleControlDisabled "data3", true
End Sub

' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  Button1.disabled = false
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
<%If ShowInBrowserWin then%>
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
<%Else%>
  Window.close   
<%End If%>
End Sub

' Intoarce in format CSV ID-urile recordurilor selectate si afiseaza un
' mesaj de avertizare daca nu s-a selectat nici o inregistrare
Function GetSelectedRecords
  Dim RecList
  RecList = TableGetSelected(tblMyTable)
  if RecList = "" then  MsgBox "You need to select at least one record.", vbOkOnly+vbExclamation
  GetSelectedRecords = RecList
End Function

' Face comutarea intre controalele de selectie a datei
Sub CheckDateRadios
 If RadioButton1.checked then 
  SimpleControlDisabled "data1", true
  SimpleControlDisabled "data2", true
  SimpleControlDisabled "data3", true
 ElseIf RadioButton2.checked then 
  SimpleControlDisabled "data1", false
  SimpleControlDisabled "data2", true
  SimpleControlDisabled "data3", true
 ElseIf RadioButton3.checked then 
  SimpleControlDisabled "data1", true
  SimpleControlDisabled "data2", false
  SimpleControlDisabled "data3", false
 End If 
End Sub

' Intoarce true daca datele calendaristice au fost completate corect
' in caz contrar afiseaza mesaj de avertisment si intoarce false
Function IsDateSelectionCorrect
 Dim re
 
 re = false
 If RadioButton1.checked then
  re = true
 ElseIf RadioButton2.checked then 
  If not IsEmpty(GetDateFromSimpleControl("data1")) then 
    re = true
  Else
    msgbox "Invalid date.", vbOkOnly+vbExclamation, "Warning"
  End If  
 ElseIf RadioButton3.checked then
  If (not IsEmpty(GetDateFromSimpleControl("data2"))) and (not IsEmpty(GetDateFromSimpleControl("data3"))) then
    If DateDiff("d", GetDateFromSimpleControl("data2"), GetDateFromSimpleControl("data3")) >= 0 then
     re = true
    Else
     msgbox "Invalid date range.", vbOkOnly+vbExclamation, "Warning"
    End If
  Else
    msgbox "Invalid date.", vbOkOnly+vbExclamation, "Warning"
  End If
 End If 
 IsDateSelectionCorrect = re
End Function

' Evenimentul apare la apasarea butonului Show results
Sub Button1_OnClick
 Dim RecList, ds
 Dim dt, dt1, dt2, dt3
 RecList=GetSelectedRecords
 if RecList="" then Exit Sub

 If RadioButton1.checked then
  ds = ""
 ElseIf RadioButton2.checked then
  If IsDateSelectionCorrect then 
    dt  = GetDateFromSimpleControl("data1")
    dt1 = CStr(Day(dt))
    dt2 = CStr(Month(dt))
    dt3 = CStr(Year(dt))
    ds = "&d1=" & dt1 & "." & dt2 & "." & dt3
  Else 
    Exit sub
  End If  
 ElseIf RadioButton3.checked then 
  If IsDateSelectionCorrect then 
    dt  = GetDateFromSimpleControl("data2")
    dt1 = CStr(Day(dt))
    dt2 = CStr(Month(dt))
    dt3 = CStr(Year(dt))
    ds = "&d1=" & dt1 & "." & dt2 & "." & dt3
    dt  = GetDateFromSimpleControl("data3")
    dt1 = CStr(Day(dt))
    dt2 = CStr(Month(dt))
    dt3 = CStr(Year(dt))
    ds = ds & "&d2=" & dt1 & "." & dt2 & "." & dt3
  Else 
    Exit sub
  End If  
 End If
  
 ShowModalDialog "studtstviewrez.asp?userid=<%=UserID%>&tstids=" & RecList & ds,, "dialogWidth=620px;dialogHeight=520px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button2_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
