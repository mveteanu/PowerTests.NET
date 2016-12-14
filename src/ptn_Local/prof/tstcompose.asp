<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/problems.asp" -->
<!-- #include file="../_serverscripts/tests.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<!-- #include file="../_serverscripts/TabControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim tstWindTitle
Dim CursID, TestID, cn, tst1
Dim tstNume, tstComentarii, tstTimp, tstSustineri, tstRandom, tstPublic
Dim tstListBoxPbsAvailable, tstListBoxPbsInclude, tstListBoxPbsExclude, tstListViewText
Dim AllowToModifyTest

AllowToModifyTest = true
CursID  = Session("CursID")
Set cn  = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")

If Request.QueryString.Count = 0 then
  tstWindTitle  = "Create new test"
  tstNume       = ""
  tstComentarii = ""
  tstTimp       = 60
  tstSustineri  = 0
  tstPublic     = true
  tstRandom     = 10
  tstListBoxPbsAvailable = GetFillSelectFromDict(GetTotalPBList(CursID, cn), "")
  tstListBoxPbsInclude   = ""
  tstListBoxPbsExclude   = ""
  tstListViewText        = ""
Else
  TestID = Request.QueryString("TstId")
  tstWindTitle  = "Edit test"
  set tst1 = New PTNTestDefinition
  tst1.LoadTestDefinition TestID, cn
  tstNume       = tst1.Name
  tstComentarii = tst1.Comments
  tstTimp       = tst1.Time
  tstSustineri  = tst1.MaxSustineri
  tstPublic     = tst1.TstPublic
  tstRandom     = tst1.MaxRandom
  tstListBoxPbsInclude   = GetFillSelectFromDict(tst1.PBIncluse, "")
  tstListBoxPbsExclude   = GetFillSelectFromDict(tst1.PBExcluse, "")
  tstListBoxPbsAvailable = GetFillSelectFromDict(GetDictDifference(GetDictDifference(GetTotalPBList(CursID, cn), tst1.PBIncluse), tst1.PBExcluse), "")
  tstListViewText        = GetFillLViewFromDict(tst1.CategFilt)
  set tst1 = nothing
  If GetTestNrSustineri(TestID, cn) > 0 then AllowToModifyTest = false
End If

cn.Close
Set cn = nothing

PrintWindow

' Intoarce codul HTML corespunzator pentru umplerea SELECT-ului
' de tipul ListView (notatie VMA) folosind taguri <OPTION> 
' cu umplere egala si un obiect Dictionary special...
Function GetFillLViewFromDict(Dict)
 Const OptionMach = "<OPTION value='@1'>@2</OPTION>"
 Dim re, it, ti1, ti2, ti3, tit

 For each it in Dict.keys
  ti1 = GetFixText(Dict.item(it).CategName, 29, "&nbsp;")
  ti2 = GetFixText(ComparationIDToString(Dict.item(it).Comparation), 20, "&nbsp;")
  ti3 = Dict.item(it).NrProblems
  tit = ti1 & "&nbsp;" & ti2 & ti3
  re  = re & Replace(Replace(OptionMach, "@2", tit), "@1", Dict.item(it).CategID) & vbCrLf
 Next
 
 GetFillLViewFromDict = re
End Function

' Converteste codul comparatiei intr-un sir cu numele sau
Function ComparationIDToString(c)
 Dim re
 Select Case c
   case 0 re = "Fix"
   case 1 re = "Min"
   case 2 re = "Max"
 End Select
 ComparationIDToString = re
End Function

Sub PrintWindow
%>  
<html>
<head>
  <title><%=tstWindTitle%></title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
  <script language=javascript src="../_clientscripts/tabControlEvents.js"></script>
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
<%OpenTabControl 8, 8, 630, 286, Array("General","Questions","Random"), 1, "PageControl1"%>

<%OpenTabContent%>

<fieldset class=TGroupBox style="width:281px;height:233px;"
          style="left:8px;top:8px;">
<legend>Test info</legend>
<span class=TLabel style="width:105px;height:13px;"
      style="left:8px;top:32px;">
Name:
</span>
<span class=TLabel style="width:105px;height:13px;"
      style="left:8px;top:72px;">
Comments:
</span>
<input id="EditNumeTst" type=text maxlength=50
       class=TEdit style="width:193px;height:21px;"
       style="left:72px;top:32px;"
       value="<%=tstNume%>">
<textarea id="EditComentsTst"
       class=TEdit style="width:257px;height:137px;"
       style="left:8px;top:88px;"><%=tstComentarii%></textarea>
</fieldset>

<fieldset class=TGroupBox style="width:313px;height:161px;"
          style="left:296px;top:8px;">
<legend>Solving rules</legend>
<span class=TLabel style="width:72px;height:13px;"
      style="left:16px;top:32px;">
Time:
</span>
<span class=TLabel style="width:93px;height:13px;"
      style="left:16px;top:104px;">
Maximum solvings:
</span>

<input id="TimpRadio1" name=radiogrp1 type=radio onclick="vbscript:HandleRadioEditDisable 1" <%=CompareToString((tstTimp = 0),false,"CHECKED","")%> class=TButton style="left:128px;top:56px;">
<input id="TimpRadio2" name=radiogrp1 type=radio onclick="vbscript:HandleRadioEditDisable 1" <%=CompareToString((tstTimp = 0),true,"CHECKED","")%> class=TButton style="left:128px;top:32px;">
<label for=TimpRadio1 class=TLabel style="width:80px;height:13px;" style="cursor:hand; left:150px;top:60px;">Limited (min)</label>
<label for=TimpRadio2 class=TLabel style="width:60px;height:13px;" style="cursor:hand; left:150px;top:36px;">Unlimited</label>
<input id="EditTimp" type=text <%=CompareToString((tstTimp = 0),true,"DISABLED","")%> maxlength=4 class=TEdit style="width:73px;height:21px;" style="left:224px;top:56px;"
       value="<%=CompareToString((tstTimp = 0),true,"",CStr(tstTimp))%>">

<input id="SustinereRadio1" name=radiogrp2 type=radio onclick="vbscript:HandleRadioEditDisable 2" <%=CompareToString((tstSustineri = 0),false,"CHECKED","")%> class=TButton style="left:128px;top:128px;">
<input id="SustinereRadio2" name=radiogrp2 type=radio onclick="vbscript:HandleRadioEditDisable 2" <%=CompareToString((tstSustineri = 0),true,"CHECKED","")%> class=TButton style="left:128px;top:104px;">
<label for=SustinereRadio1 class=TLabel style="width:60px;height:13px;" style="cursor:hand; left:150px;top:132px;">Limited</label>
<label for=SustinereRadio2 class=TLabel style="width:60px;height:13px;" style="cursor:hand; left:150px;top:108px;">Unlimited</label>
<input id="EditSustinere" type=text <%=CompareToString((tstSustineri = 0),true,"DISABLED","")%> maxlength=4 class=TEdit style="width:73px;height:21px;" style="left:224px;top:128px;"
       value="<%=CompareToString((tstSustineri = 0),true,"",CStr(tstSustineri))%>">
</fieldset>

<fieldset class=TGroupBox style="width:313px;height:65px;"
          style="left:296px;top:176px;">
<legend>Misc</legend>
<span class=TLabel style="width:107px;height:13px;"
      style="left:16px;top:16px;">
Visible to students:
</span>
<input id="PublicRadio1" <%=CompareToString(tstPublic,true,"CHECKED","")%> name=radiogrp3 type=radio class=TButton style="left:160px;top:16px;">
<input id="PublicRadio2" <%=CompareToString(tstPublic,false,"CHECKED","")%> name=radiogrp3 type=radio class=TButton style="left:160px;top:40px;">
<label for=PublicRadio1 class=TLabel style="width:24px;height:13px;" style="cursor:hand; left:182px;top:20px;">Yes</label>
<label for=PublicRadio2 class=TLabel style="width:24px;height:13px;" style="cursor:hand; left:182px;top:44px;">No</label>
</fieldset>

<%CloseTabContent%>

<%OpenTabContent%>
<span class=TLabel style="width:145px;height:13px;"
      style="left:16px;top:8px;">
Include these questions:
</span>
<span class=TLabel style="width:145px;height:13px;"
      style="left:230px;top:8px;">
Available questions:
</span>
<span class=TLabel style="width:145px;height:13px;"
      style="left:440px;top:8px;">
Don't include questions:
</span>

<a     class=TLinkLabel href=# title="Preview al problemelor selectate" style="left:16px;top:232px;"
       onclick="vbscript:HandlePreviewProblems ListBox1">View selected</a>
<a     class=TLinkLabel href=# title="Preview al problemelor selectate" style="left:230px;top:232px;"
       onclick="vbscript:HandlePreviewProblems ListBox2">View selected</a>
<a     class=TLinkLabel href=# title="Preview al problemelor selectate" style="left:440px;top:232px;"
       onclick="vbscript:HandlePreviewProblems ListBox3">View selected</a>

<select id="ListBox1" size=2 multiple
        class=TListBox style="width:161px;height:201px;"
        style="left:16px;top:24px;">
<%=tstListBoxPbsInclude%>
</select>

<select id="ListBox2" size=2 multiple
        class=TListBox style="width:161px;height:201px;"
        style="left:230px;top:24px;">
<%=tstListBoxPbsAvailable%>
</select>

<select id="ListBox3" size=2 multiple
        class=TListBox style="width:161px;height:201px;"
        style="left:440px;top:24px;">
<%=tstListBoxPbsExclude%>
</select>

<input type=button value="<"
       class=TButton style="width:25px;height:25px;"
       style="left:190px;top:56px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, "<"'>
<input type=button value="<<"
       class=TButton style="width:25px;height:25px;"
       style="left:190px;top:88px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, "<<"'>
<input type=button value=">"
       class=TButton style="width:25px;height:25px;"
       style="left:190px;top:136px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, ">"'>
<input type=button value=">>"
       class=TButton style="width:25px;height:25px;"
       style="left:190px;top:168px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox1, ListBox2, ">>"'>

<input type=button value=">"
       class=TButton style="width:25px;height:25px;"
       style="left:404px;top:56px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox2, ListBox3, ">"'>
<input type=button value=">>"
       class=TButton style="width:25px;height:25px;"
       style="left:404px;top:88px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox2, ListBox3, ">>"'>
<input type=button value="<"
       class=TButton style="width:25px;height:25px;"
       style="left:404px;top:136px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox2, ListBox3, "<"'>
<input type=button value="<<"
       class=TButton style="width:25px;height:25px;"
       style="left:404px;top:168px;"
       onclick='vbscript:HandleButtonsBetweenSelects ListBox2, ListBox3, "<<"'>
<%CloseTabContent%>

<%OpenTabContent%>
<span class=TLabel style="width:200px;height:13px;"
      style="left:12px;top:16px;">
Maximum number of random questions:
</span>
<input id="EditMaxRnd" type=text maxlength=50
       class=TEdit style="width:121px;height:21px;"
       style="left:12px;top:32px;"
       value="<%=CStr(tstRandom)%>">
<span class=TLabel style="width:127px;height:13px;"
      style="left:12px;top:64px;">
Random selection criteria:
</span>

<div class=TForm style="width:510px;height:162px;left:12px;top:80px;">
<span class=TLabel unselectable="on"
      style="width:505px;height:18px;font-weight:bold;font-size:10pt;font-family:'Courier New';" 
      style="border:groove thin;left:1px;top:0px;">
<%=GetFixText("Category", 30, "&nbsp;") & GetFixText("Comparation", 20, "&nbsp;") & "Questions"%>
</span>
<select id="ListView1" size=2
        class=TListBox style="width:505px;height:140px;" style="left:1px;top:18px;"
        style="font-size:10pt;font-family:'Courier New';">
<%=tstListViewText%>        
</select>
</div>

<input id="ListViewAdd" type=button value="Add criteria" title="Add random selection criteria"
       class=TButton style="width:75px;height:25px;"
       style="left:533px;top:80px;">
<input id="ListViewDel" type=button value="Del criteria" title="Delete criteria"
       class=TButton style="width:75px;height:25px;"
       style="left:533px;top:120px;">
<input id="ListViewEdt" type=button value="Edit" title="Edit random selection criteria"
       class=TButton style="width:75px;height:25px;"
       style="left:533px;top:160px;">

<%CloseTabContent%>


<%CloseTabControl%>

<input id="Button1" type=button value="Save" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:239px;top:300px;">
<input id="Button2" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:327px;top:300px;">
</div>


<script language=vbscript src="../_clientscripts/MiscControlUtils.vbs"></script>
<script language=vbscript src="../_clientscripts/SelectControlUtils.vbs"></script>
<script language=vbscript src="../_clientscripts/LBLVControlUtils.vbs"></script>
<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Gestioneaza comutarea intre RadioButton-ele de tip Limitat/Nelimitat
' si EditBox-ul corespunzator pozitiei Limitat
Sub HandleRadioEditDisable(sn)
 select case sn
   case 1 TwoRadioOneEditDisable TimpRadio2, EditTimp
   case 2 TwoRadioOneEditDisable SustinereRadio2, EditSustinere
 end select
End Sub

' Este folosita de butoanele de tip link ce permit vizualizarea
' problemelor selectate din SELECT-urile cu probleme
Sub HandlePreviewProblems(objListBox)
  PreviewProblems GetItemsFromSelect(objListBox, true)
End Sub

' Deschide o fereastra de preview cu problemele specificate prin 
' lista de ID-uri PBIDs
Sub PreviewProblems(PBIDs)
  If PBIDs<>"" then
   ShowModalDialog "pbviewscr.asp?PBIDs=" & PBIDs, "", "dialogWidth=720px;dialogHeight=500px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
  Else
   MsgBox "You need to select at least one question.", vbOkOnly+vbExclamation
  End If 
End Sub

' Trateaza evenimentul aparut la apasarea butonului ListView Del
Sub ListViewDel_onclick
  Dim selitem
  
  selitem = ListView1.selectedIndex 
  If selitem<>-1 then
    ListView1.Options.Remove(selitem)
  Else
    MsgBox "You need to select a record.", vbOkOnly+vbExclamation
  End If
End Sub


' Trateaza evenimentul aparut la apasarea butonului ListView Add
Sub ListViewAdd_onclick
  Dim CategFilt
  Dim rpar, rp1, rp2, rp3, newel
  
  CategFilt = ShowModalDialog("tsteditfilter.asp", , "dialogWidth=337px;dialogHeight=217px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If CategFilt<>"" then
    rpar = Split(CategFilt, "|", -1, 1)
    rp1  = GetFixText(rpar(3), 29, "&nbsp;")
    rp2  = GetFixText(Array("Fix","Min","Max")(CInt(rpar(1))), 20, "&nbsp;")
    rp3  = rpar(2)
    
    Set newel = document.createElement("OPTION")
    ListView1.Options.Add newel
    newel.innerHTML = rp1 & "&nbsp;" & rp2 & rp3
    newel.Value     = rpar(0)
    Set newel = nothing
  End If
End Sub


' Trateaza evenimentul aparut la apasarea butonului ListView Edt
Sub ListViewEdt_onclick
  Dim CategFilt, selitem
  Dim par, p1, p2, p3, pqs
  Dim rpar, rp1, rp2, rp3
  
  selitem = ListView1.selectedIndex 
  If selitem=-1 then
    MsgBox "You need to select a record.", vbOkOnly+vbExclamation
    Exit Sub
  End If
   
  p1 = ListView1.Options(selitem).Value
  par = NonvidSplit(ListView1.options(selitem).innerHTML, "&nbsp;")
  Select Case LCase(par(1))
   case "fix" p2 = 0
   case "min" p2 = 1
   case "max" p2 = 2
  End Select
  p3 = par(2)
  pqs = "p1=" & CStr(p1) & "&p2=" & CStr(p2) & "&p3=" & CStr(p3)
  
  CategFilt = ShowModalDialog("tsteditfilter.asp?" & pqs, , "dialogWidth=337px;dialogHeight=217px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If CategFilt<>"" then
    rpar = Split(CategFilt, "|", -1, 1)
    rp1  = GetFixText(rpar(3), 29, "&nbsp;")
    rp2  = GetFixText(Array("Fix","Min","Max")(CInt(rpar(1))), 20, "&nbsp;")
    rp3  = rpar(2)

    ListView1.options(selitem).Value     = rpar(0)
    ListView1.options(selitem).innerHTML = rp1 & "&nbsp;" & rp2 & rp3
  End If
End Sub


' Impacheteaza datele din fereastra intr-un string. Datele
' principale sunt separate prin caracterul Chr(1)
' Daca regulile de definire nu sunt corecte functia intoarce 
' sirul vid
Function PackContolsData
 Dim locNumetest, locComentarii, locTimp, locSustineri, locRandom, locPublic 
 Dim locIncPb, locExcPb, locCatFilt
 Dim re, it, par, p2, pqs
 
 locNumetest   = EditNumeTst.Value
 locComentarii = EditComentsTst.value 
 If TimpRadio2.Checked then locTimp = 0 Else locTimp = EditTimp.Value
 If SustinereRadio2.Checked then locSustineri = 0 Else locSustineri = EditSustinere.Value 
 locPublic = LCase(CStr(PublicRadio1.Checked))
 locRandom = EditMaxRnd.Value
 For each it in ListBox1.options 
  locIncPb = locIncPb & it.Value & ","
 Next
 If locIncPb<>"" then locIncPb = Left(locIncPb, Len(locIncPb)-Len(","))
 For each it in ListBox3.options 
  locExcPb = locExcPb & it.Value & "," 
 Next
 If locExcPb<>"" then locExcPb = Left(locExcPb, Len(locExcPb)-Len(","))
 For each it in ListView1.options 
  par = NonvidSplit(it.innerHTML, "&nbsp;")
  Select Case LCase(par(1))
   case "fix" p2 = 0
   case "min" p2 = 1
   case "max" p2 = 2
  End Select
  pqs = CStr(it.Value) & "," & CStr(p2) & "," & CStr(par(2))
  locCatFilt = locCatFilt & pqs & "|"
 Next
 If locCatFilt<>"" then locCatFilt = Left(locCatFilt, Len(locCatFilt)-Len("|"))
 
 If (locNumetest="") or _ 
    (not IsNumeric(locTimp)) or _
    (not IsNumeric(locSustineri)) or _
    ((locIncPb="") and (not IsNumeric(locRandom))) Then
   re = ""
 Else 
   re = locNumetest & Chr(1) &_
        locComentarii & Chr(1) &_
        locTimp & Chr(1) &_ 
        locSustineri & Chr(1) &_
        locRandom & Chr(1) &_ 
        locPublic & Chr(1) &_
        locIncPb & Chr(1) &_
        locExcPb & Chr(1) &_
        locCatFilt
 End If  
 
 PackContolsData = re
End Function

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
<%If AllowToModifyTest then%>
  Dim re
  
  re = PackContolsData
  If re = "" then
    msgbox "You need to enter required data.", vbOkOnly+vbExclamation
  Else
    Window.returnValue = re
    Window.close 
  End If
<%Else%>
  msgbox "This test cannot be edited at this moment," & vbCrLf & "because was allready solved by some students.", vbOkOnly+vbExclamation
<%End If%>
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.returnValue = ""
  Window.close 
End Sub
</script>

</body>
</html>
<%
End Sub ' PrintWindow
%>


