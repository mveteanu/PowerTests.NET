<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/users.asp" -->
<!-- #include file="../_serverscripts/TableControl.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<script language=vbscript src="../_clientscripts/TableControlEvents.vbs"></script>
<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:535px;height:265px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<div id="Form1" style="visibility:hidden;"
     class=TForm style="width:535px;height:265px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 12, 8, 200, Array("Language", "Users", "Default"), Array(305,100,100), 1, "languageslist_dat.asp", true, "MyTable"%>
<input DISABLED id="Button1" type=button value="Add" title="Adauga o noua limba in aplicatie"
       class=TButton style="width:85px;height:25px;"
       style="left:12px;top:223px;">
<input DISABLED id="Button2" type=button value="Delete" title="Delete selected language"
       class=TButton style="width:85px;height:25px;"
       style="left:117px;top:223px;">
<input id="Button3" type=button value="Rename" title="Rename selected language"
       class=TButton style="width:85px;height:25px;"
       style="left:222px;top:223px;">
<input id="Button4" type=button value="Set as default" title="Set selected language as default system language"
       class=TButton style="width:85px;height:25px;"
       style="left:327px;top:223px;">
<input id="Button5" type=button value="Close" title="Close form"
       class=TButton style="width:85px;height:25px;"
       style="left:432px;top:223px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="languageslist_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
' Activeaza/dezactiveaza butoanele
Sub ActivateButtons(state)
	For i = 1 to 5
		Form1.All("Button" & CStr(i)).disabled = not state
	Next
End Sub

' Apare la incarcarea documentului
Sub window_onload
	Form1.style.visibility = "visible"
	WaitforForm.style.visibility = "hidden"
End Sub

' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
	ActivateButtons true
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
	tdcMyTable.DataURL = tdcMyTable.DataURL
	tdcMyTable.Reset
End Sub

' Intoarce in format CSV ID-urile recordurilor selectate si afiseaza un
' mesaj de avertizare daca nu s-a selectat nici o inregistrare
Function GetSelectedRecords
	Dim RecList
	RecList = TableGetSelected(tblMyTable)
	If Len(RecList) = 0 Then  MsgBox "You need to select at least one record.", vbOkOnly+vbExclamation
	GetSelectedRecords = RecList
End Function

' Intoarce sub forma de string ID-ul recordului selectat.
' Daca nu se selecteaza nici o inregistrare sau se selecteaza mai mult de una
' atunci se afiseaza un mesaj si se intoarce sirul vid.
Function GetSelectedRecord
	Dim RecList
	Dim RecArray
  
	RecList  = TableGetSelected(tblMyTable)
	RecArray = Split(RecList,",",-1,1)
	If (UBound(RecArray)-LBound(RecArray))<>0 then 
		MsgBox "You need to select at least one record.", vbOkOnly+vbExclamation
		RecList = ""
	End If  
	GetSelectedRecord = RecList
End Function

' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
	WaitforForm.style.visibility = "hidden"
	Form1.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Add new language
Sub Button1_OnClick
	DateLang = ShowModalDialog("languageadd_cli.asp",, "dialogWidth=334px;dialogHeight=243px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
	If Not IsArray(DateLang) Then Exit Sub

	FormularH.SelectAction.value = "add"
	FormularH.SelectList.value	= DateLang(1) ' based on - 0 if is a completely new language
	FormularH.SelectValue.value = DateLang(0) ' lang name
	ActivateButtons false
	FormularH.submit 
End Sub


' Evenimentul apare la apasarea butonului Delete Language
Sub Button2_OnClick
	' Aici tre sa cer perminiunea sa stearga limba selectata..
	' Daca limba este folosita de mai multi utilizatori sa 
	' intrebe ce alta limba li se va atribui acelor utilizatori
	Dim NrUsers
	Dim bPref
	Dim NewLang

	RecList = GetSelectedRecord
	If Len(RecList)=0 then Exit Sub
	
	NrUsers = CLng(GetTDCData(tdcMyTable, RecList, Array("Users"))(0))
	bPref = CBool(GetTDCData(tdcMyTable, RecList, Array("bpreferata"))(0))
	If bPref Then
		MsgBox "You cannot delete application default language." & vbCrLf & "Change first the default language and try again.", vbOkOnly+vbExclamation
		Exit Sub
	End If
	
	If NrUsers > 0 Then
		NewLang = ShowModalDialog("languagedel_cli.asp?id=" & RecList & "&nru=" & NrUsers, , "dialogWidth=280px;dialogHeight=170px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
		If Len(NewLang) = 0 Then Exit Sub
	Else
		NewLang = 0
	End If
	
	If MsgBox("Are you sure you want to delete the selected language?", vbYesNo+vbQuestion, "Confirm") = vbYes Then
		FormularH.SelectAction.value = "del"
		FormularH.SelectList.value	= RecList
		FormularH.SelectValue.value = NewLang
		ActivateButtons false
		FormularH.submit 
	End If
End Sub


' Evenimentul apare la apasarea butonului Rename Language
Sub Button3_OnClick
	Dim RecList, LangNewName
	RecList = GetSelectedRecord
	If Len(RecList)=0 then Exit Sub
  
	LangNewName = ShowModalDialog("languageren_cli.asp?id=" & RecList, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
	If Len(LangNewName)<>0 Then
		FormularH.SelectAction.value = "ren"
		FormularH.SelectList.value	= RecList
		FormularH.SelectValue.value = LangNewName
		ActivateButtons false
		FormularH.submit 
	End If 
End Sub


' Evenimentul apare la apasarea butonului Preferred language
Sub Button4_OnClick
	Dim RecList, LangNewName
	Dim bPref
	RecList = GetSelectedRecord
	If Len(RecList)=0 then Exit Sub
  
	bPref = CBool(GetTDCData(tdcMyTable, RecList, Array("bpreferata"))(0))
	If bPref Then
		MsgBox "Language is allready set as default." & vbCrLf & "Select other language that you want to set as default.", vbOkOnly+vbExclamation
	Else
		FormularH.SelectAction.value = "preflang"
		FormularH.SelectList.value	= RecList
		ActivateButtons false
		FormularH.submit 
	End If
End Sub

' Evenimentul apare la apasarea butonului Close
Sub Button5_OnClick
	HideAllDivs
End Sub
</script>

</BODY>
</HTML>
