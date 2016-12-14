<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<!-- #include file="../_serverscripts/tests.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim CategsList

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
CategsList = GetFillSelectFromArrDict(GetTstCategsDict(Session("CursID") ,cn))
cn.Close
set cn = nothing

Function GetTstCategsDict(cursid, objCon)
 Dim re, catid, catdict, catstr, alltst, alltststr, hidtst, astra
 Dim myCmd, rs

 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetTstCategByCursID"
 myCmd.CommandType = adCmdStoredProc
 Set rs  = myCmd.Execute(,CLng(cursid))
 Set re = CreateObject("Scripting.Dictionary")

 re.Add "alltst", Array("","")
 re.Add "uncat", Array("","")
 
 ' Face dictionarul hidtst cu testele care mostenesc atributul de
 ' hidden de la o categorie ascunsa in care se afla
 set hidtst = CreateObject("Scripting.Dictionary")
 do until rs.EOF
   If not rs.Fields("categtstpublica").value then
    catid = rs.Fields("id_categtst").value
    set hidtst = GetDictSum(Array(hidtst,GetCategTstList(catid, true, objCon)))
   End If
   rs.MoveNext
 loop
 If not rs.BOF then rs.MoveFirst
 
 ' Se adauga categoriile publice care contin cel putin un test
 ' cu testele lor publice dar fara testele care sunt ascunse
 ' prin mostenire de la o alta categorie
 set alltst = GetTotalTstList(cursid, true, objCon)
 do until rs.EOF
   If rs.Fields("categtstpublica").value then 
     catid = rs.Fields("id_categtst").value
     set catdict = GetDictDifference(GetCategTstList(catid, true, objCon), hidtst)
     catstr = FilterStringFromDict(catdict)
     set alltst = GetDictDifference(alltst, catdict)
     set catdict = nothing
     If catstr<>"" then re.Add CStr(catid), Array(rs.Fields("numecateg").value, catstr)
   End If
   rs.MoveNext
 loop

 rs.Close
 Set rs = nothing

 ' Creaza si pseudocategoriile:
 '   - teste necatalogate
 '   - toate testele
 alltststr = FilterStringFromDict(GetDictDifference(alltst, hidtst))
 If alltststr="" then 
   re.Remove("uncat") 
 Else 
   re.Item("uncat") = Array("Uncategorised tests", alltststr)
 End If  
 
 For Each Ita in re.Keys 
  If re.Item(Ita)(1)<>"" then astra = astra & re.Item(Ita)(1) & "|"
 Next
 If astra<>"" then 
   astra = Left(astra, Len(astra)-Len("|"))
   re.Item("alltst") = Array("All tests",astra)
 Else
   re.Remove("alltst")
 End If  

 Set GetTstCategsDict = re
End Function

Function FilterStringFromDict(dict)
 Dim re
 For each it in dict.keys
  re = re & "id=" & CStr(it) & "|"
 Next
 If re<>"" then re = Left(re, Len(re)-Len("|"))
 FilterStringFromDict = re
End Function

Function GetFillSelectFromArrDict(Dict)
 Const OptionMach = "<OPTION value='@1'>@2</OPTION>"
 Dim re, re1, it

 For each it in Dict.keys
  re  = re & Replace(Replace(OptionMach, "@2", Dict.item(it)(0)), "@1", Dict.item(it)(1)) & vbCrLf
 Next
 
 GetFillSelectFromArrDict = re
End Function
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:610px;height:308px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>


<div id="Form1" style="visibility:hidden;" unselectable="on"
     class=TForm style="width:610px;height:308px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<div class=TForm style="width:585px;height:41px;" unselectable="on"
     style="left:8px;top:8px;">
<select id="ComboBox1"
        class=TComboBox style="width:560px;left:8px;top:10px;">
<option SELECTED value="id=-1">Select test category</option>
<%=CategsList%>
</select>
</div>
<%CreateTableControl 8, 48, 217, Array("Test name","Time","Solvings","Maximum solvings"), Array(195,130,130,130), 1, "studseltst_dat.asp" , true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Test info" title="Information about test"
       class=TButton style="width:85px;height:25px;"
       style="left:151px;top:272px;">
<input DISABLED id="Button2" type=button value="Take test" title="Take selected test"
       class=TButton style="width:85px;height:25px;"
       style="left:259px;top:272px;">
<input id="Button3" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:367px;top:272px;">
</div>

<script language=vbscript>
Dim TDCFirtTime

' Schimba starea activ/inactiv a butoanelor
Sub ActivateButtons(btnsstate)
 Dim i
 
 for i=1 to 2
   Form1.all("Button"&CStr(i)).disabled = not btnsstate
 next  
End Sub

' Apare la incarcarea documentului
Sub window_onload
   Form1.style.visibility = "visible"
   WaitforForm.style.visibility = "hidden"
End Sub

' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  If IsEmpty(TDCFirtTime) then
    TDCFirtTime = false
    tdcMyTable.filter = "id=-1" 
    tdcMyTable.reset
  End If  
    ActivateButtons true
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub

' Apare la selectarea din ComboBox
Sub ComboBox1_OnChange
  tdcMyTable.filter = ComboBox1.value 
  tdcMyTable.reset
End Sub

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
  WaitforForm.style.visibility = "hidden"
  Form1.style.visibility = "hidden"
End Sub

' Intoarce sub forma de string ID-ul recordului selectat.
' Daca nu se selecteaza nici o inregistrare sau se selecteaza mai mult de una
' atunci se afiseaza un mesaj si se intoarce sirul vid.
Function GetSelectedRecord
  Dim RecList
  Dim RecArray
  
  RecList  = TableGetSelected(tblMyTable)
  RecArray = Split(RecList,",",-1,1)
  If (UBound(RecArray)-LBound(RecArray))<>0 then 
    MsgBox "You need to select a record.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function

' Evenimentul apare la apasarea butonului Info test
Sub Button1_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  ShowModalDialog "studtstinfo.asp?TestID=" & RecList,, "dialogWidth=530px;dialogHeight=278px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub  

' Intoarce false daca nu se permite rularea unui anumit test
' datorita atingerii numarului maxim de sustineri permis
Function AllowToRunTest(idtest)
 Dim tar, sustineri, maxsustineri
 Dim re
 
 tar = GetTDCData(tdcMyTable, idtest, Array("Solvings","Maximum solvings"))
 sustineri = tar(0) : maxsustineri = tar(1)
 If Not IsNumeric(maxsustineri) Then
   re = true
 ElseIf CInt(sustineri) >= CInt(maxsustineri) Then  
   re = false
 Else
   re = true  
 End If
 AllowToRunTest = re
End Function

' Evenimentul apare la apasarea butonului Rezolva testul
Sub Button2_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  If not AllowToRunTest(RecList) then
    msgbox "You cannot take selected test because " & VbCrLf & "you reached maximum number of solvings.", vbOkOnly+vbExclamation
  else
    ShowModalDialog "studtstview.asp?TstID=" & RecList, , "dialogWidth=720px;dialogHeight=560px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; resizable=no; status=no"
    ReloadTDC
  End If  
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button3_OnClick
  HideAllDivs
End Sub
</script>

</BODY>
</HTML>
