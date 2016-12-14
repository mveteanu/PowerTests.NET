<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/ControlUtils.asp" -->
<!-- #include file="../_serverscripts/settings.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1
	
	' Insereaza 2 controale TDC pentru gestionarea "Paginilor" si a "Controalelor"
	Sub InsertTDCControls
		Call AddTDC("TDCPage", Chr(1), "languagetexts_dat.asp?w=1")
		Response.Write "<span style='display:none;' DATASRC=#TDCPage DATAFLD='id'></span>"

		Call AddTDC("TDCControl", Chr(1), "languagetexts_dat.asp?w=2")
		Response.Write "<span style='display:none;' DATASRC=#TDCControl DATAFLD='id'></span>"

		Call AddTDC("TDCTexts", Chr(1), "languagetexts_dat.asp?w=3&idc=-1")
		Response.Write "<span style='display:none;' DATASRC=#TDCTexts DATAFLD='id'></span>"
	End Sub


	' Intoarce sub forma de string CSV id-urile limbilor definite
	Function GetLanguagesIds
		Const separator = ", "
		Dim db, rs
		Dim idprefered
		Dim s
		
		s = ""
		Set db = New clsDatabase
		Set rs = db.GetRs("SELECT id FROM TBLanguages", Empty)
		Do While Not rs.EOF
			s = s & rs.Fields("id").Value & separator
			rs.MoveNext
		Loop
		Set rs = Nothing
		Set db = Nothing
		
		If Len(s) > 0 Then
			s = Left(s, Len(s) - Len(separator))
		End If

		GetLanguagesIds = s
	End Function


	' Insereaza casutele de tip TEXTAREA ce contin textele ce se vor traduce
	Sub InsertLanguageBoxes
		Dim db, rs
		Dim i, id, name
		Dim idprefered, p
		
		Set db = New clsDatabase
		idprefered = GetPreferredLanguage(db.Connection)
		i = 1
		
		Set rs = db.GetRs("SELECT id, langname FROM TBLanguages", Empty)
		Do While Not rs.EOF
			id = rs.Fields("id").Value
			name = rs.Fields("langname").Value
			If id = idprefered Then 
				p = 0 
			Else 
				p = i
				i = i + 1
			End If
			Response.Write "<span class=TLabel style='width:300px;height:13px;left:8px;top:" & CStr((64 * p) + 8) & "px;' unselectable='on'>" & name & "</span>"
			Response.Write "<textarea disabled name='txt_" & id & "' id='txt_" & id & "' class=TEdit style='left:8px;top:" & CStr((64 * p) + 24) & "px;width:337px;height:41px;'></textarea>" & vbCrLf
			rs.MoveNext
		Loop
		Set rs = Nothing
		Set db = Nothing
	End Sub
%>
<HTML>
<head>
 <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<%Call InsertTDCControls%>

<div id="WaitforForm" style="visibility:visible;"
     class=TForm style="width:604px;height:364px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<div id="Form1" style="visibility:hidden;" unselectable="on"
     class=TForm style="width:604px;height:364px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">

<div class=TForm style="width:233px;height:321px;" unselectable="on"
     style="left:0px;top:0px;">
<span class=TLabel style="width:200px;height:13px;" unselectable="on"
      style="left:8px;top:8px;">
Select form
</span>
<select id="ComboBoxPage"
        class=TComboBox style="width:217px;left:8px;top:24px;">
</select>
<span class=TLabel style="width:200px;height:13px;" unselectable="on"
      style="left:8px;top:56px;">
Select control
</span>
<select id="ListBoxControl" 
        size=2
        class=TListBox style="width:217px;height:233px;"
        style="left:8px;top:72px;">
</select>

</div>

<div id="LangTexts" class=TForm style="width:368px;height:273px;" unselectable="on"
     style="left:232px;top:0px;">
	<%Call InsertLanguageBoxes%>     
</div>


<div id="LangButtons" class=TForm style="width:368px;height:49px;" unselectable="on"
     style="left:232px;top:272px;">
<input id="ButtonEdit" type=button value="Edit" title="Edit texts"
       class=TButton style="width:75px;height:25px;"
       style="left:8px;top:8px;">
<input disabled id="ButtonEditOK" type=button value="OK" title="Save changes"
       class=TButton style="width:75px;height:25px;"
       style="left:200px;top:8px;">
<input disabled id="ButtonEditCancel" type=button value="Cancel" title="Cancel changes"
       class=TButton style="width:75px;height:25px;"
       style="left:280px;top:8px;">
</div>


<div id="LangWait" style="visibility:hidden;" class=TForm style="width:368px;height:321px;" unselectable="on"
     style="left:232px;top:0px;">
	<table border=0 width=100% height=100%><tr><td align=center valign=center>
	<span>Load texts...</span>
	<span id=TextLoaddingErrorMsg></span>
	</td></tr></table>
</div>


<button id="ButtonPage" title="Manage application forms"
       class=TButton style="width:104px;height:25px;"
       style="left:8px;top:328px;">Manage forms <font face="Webdings">6</font></button>
<button id="ButtonControl" title="Manage form controls"
       class=TButton style="width:104px;height:25px;"
       style="left:120px;top:328px;">Manage controls <font face="Webdings">6</font></button>
<button id="ButtonClose" title="Close form"
       class=TButton style="width:89px;height:25px;"
       style="left:504px;top:328px;">Close</button>
     
</div>


<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="languagetexts_ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>


<script language="vbscript" src="../_clientscripts/utils.vbs"></script>
<script language="vbscript" src="../_clientscripts/menu.vbs"></script>
<script language="vbscript" src="../_clientscripts/SelectControlUtils.vbs"></script>
<script language=vbscript>
Const ElemToLoad = 3
Dim MyMenuPage, MyMenuControl
Dim ElemLoaded, lngIDControl
Dim arLanguages

<%
' Insereaza o bucata de script client cu id-urile limbilor definite
slng = GetLanguagesIds()
If Len(slng) > 0 Then
	Response.Write "arLanguages = Array(" & slng & ")"
Else
	Response.Write "arLanguages = Array()"
End If
%>

MenuPageItems		= Array("Add form","Delete form", "<HR>", "Rename form")
MenuControlItems	= Array("Add control","Delete control", "<HR>", "Rename control")

ElemLoaded = 0
lngIDControl = -1

lngPageToPosition = -1
lngControlToPosition = -1

' Intoarce true cand pagina si toate TDC-urile sunt incarcate complet
Function AllElemLoaded
	ElemLoaded = ElemLoaded + 1
	AllElemLoaded = CBool(ElemLoaded = ElemToLoad)
End Function

' La incarcarea completa a documentului trebuie ascuns div-ul cu
' mesajul de asteptare si afisa div-ul cu formul principal
Sub Window_OnLoad
	If AllElemLoaded Then ShowMainDiv
End Sub


Sub TDCPage_OnDataSetComplete
	If AllElemLoaded Then ShowMainDiv
	
	' Partea asta trebuie executata doar la intoarcerea de pe server
	' in urma adaugarii/stergerii/redenumirii unei pagini
	If lngPageToPosition > -1 Then
		Call ClearTexts()
		Call FillBoxes(lngPageToPosition)
	End If
End Sub


Sub TDCControl_OnDataSetComplete
	If AllElemLoaded Then ShowMainDiv

	' Partea asta trebuie executata doar la intoarcerea de pe server
	' in urma adaugarii/stergerii/redenumirii unui control
	If lngControlToPosition > -1 Then
		Call ClearSelect(ListBoxControl)
		If ComboBoxPage.selectedIndex > -1 Then
			Call FillSelectFromFilteredRS(ListBoxControl, TDCControl.Recordset.Clone, "id", "controlname", "id_page", ComboBoxPage.value, -1)
			If lngControlToPosition = 0 Then
				If ListBoxControl.Options.Length > 0 Then 
					ListBoxControl.selectedIndex = 0
					Call ListBoxControl_OnChange
				Else
					Call ClearTexts()
				End If
			Else
				ListBoxControl.value = lngControlToPosition
				Call ListBoxControl_OnChange
			End If
		End If
	End If
End Sub


' Umple combobox-ul cu paginile si listbox-ul cu controale
Sub FillBoxes(lngDefaultPage)
	Call ClearSelect(ComboBoxPage)
	Call ClearSelect(ListBoxControl)
	Call FillSelectFromFilteredRS(ComboBoxPage, TDCPage.Recordset.Clone, "id", "name", "", -1, -1)
	If lngDefaultPage = 0 Then
		If ComboBoxPage.Options.Length > 0 Then ComboBoxPage.selectedIndex = 0
	Else
		ComboBoxPage.value = lngDefaultPage
	End If
	If ComboBoxPage.selectedIndex > -1 Then
		Call FillSelectFromFilteredRS(ListBoxControl, TDCControl.Recordset.Clone, "id", "controlname", "id_page", ComboBoxPage.value, -1)
	End If
End Sub


' Afiseaza DIV-ul principal si umple listele cu datele luate din TDC-uri
Sub ShowMainDiv
	Call FillBoxes(0)

	Form1.style.visibility = "visible"
	WaitforForm.style.visibility = "hidden"
End Sub


' Umple textbox-urile cu textul in toate limbile
' cu informatiile luate din TDC
Function FillTextsFromTDC()
	Dim id_language
	Dim textvalue
	Dim rs
	
	Call ClearTexts()
	Set rs = TDCTexts.Recordset.Clone
	If not rs.EOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			id_language = rs.Fields("id_lang")
			textvalue = rs.Fields("textvalue")
			document.all("txt_" & id_language).Value = textvalue
			rs.MoveNext
		Loop
		FillTextsFromTDC = true
	Else
		FillTextsFromTDC = false
	End If
	Set rs = Nothing
End Function


' Sterge casutele cu toate textele corespunzatoare limbilor definite
Sub ClearTexts()
	For Each id In arLanguages
		With document.all("txt_" & id)
			.Value = ""
			.Disabled = True
		End With
	Next
End Sub


' Trateaza evenimentul care apare la selectarea unei alte
' pagini din combobox-ul cu paginile de controale
Sub ComboBoxPage_OnChange
	If ComboBoxPage.selectedIndex > -1 Then
		Call ClearTexts()
		Call ClearSelect(ListBoxControl)
		Call FillSelectFromFilteredRS(ListBoxControl, TDCControl.Recordset.Clone, "id", "controlname", "id_page", ComboBoxPage.value, -1)
	End If
End Sub


' Trateaza evenimentul care apare cand se selecteaza un anumit
' control din listbox-ul cu controale
Sub ListBoxControl_OnChange
	If ListBoxControl.selectedIndex > -1 Then
		lngIDControl = ListBoxControl.Value 
		TextLoaddingErrorMsg.innerHTML = ""
		LangTexts.style.visibility = "hidden"
		LangButtons.style.visibility = "hidden"
		LangWait.style.visibility = "inherit"

		TDCTexts.DataURL = "languagetexts_dat.asp?w=3&idc=" & lngIDControl
		TDCTexts.Reset
	End If
End Sub


' Trateaza evenimentul care apare cand se incarca complet TDC-ul
' cu textele in toate limbile definite pentru controlul selectat
Sub TDCTexts_OnDataSetComplete
		If lngIDControl = -1 Then Exit Sub
		If FillTextsFromTDC() Then
			ButtonEdit.Disabled = False
			ButtonEditOK.Disabled = True
			ButtonEditCancel.Disabled = True
		
			LangTexts.style.visibility = "inherit"
			LangButtons.style.visibility = "inherit"
			LangWait.style.visibility = "hidden"
		Else
			TextLoaddingErrorMsg.innerHTML = "<br><br><font color=red>Error! Inexistent text(s)!</font>"
		End If
End Sub


' Ascunde toate div-urile. Subrutina e folosita in momentul in
' care se apasa butonul Close.
Sub HideAllDivs
	WaitforForm.style.visibility = "hidden"
	Form1.style.visibility = "hidden"
End Sub


' Evenimentul apare la apasarea butonului Page
Sub ButtonPage_OnClick
	Dim leftm, topm
   
	leftm = 2 + StyleSizeToInt(ButtonPage.style.left) + StyleSizeToInt(Form1.style.left)
	topm  = 2 + StyleSizeToInt(ButtonPage.style.top)  + StyleSizeToInt(ButtonPage.style.height) + StyleSizeToInt(Form1.style.top)
	Set MyMenuPage = ShowMenu(leftm, topm, 140, "HandleMenuPage", MenuPageItems)
End Sub


' Evenimentul apare la apasarea butonului Control
Sub ButtonControl_OnClick
	Dim leftm, topm
   
	leftm = 2 + StyleSizeToInt(ButtonControl.style.left) + StyleSizeToInt(Form1.style.left)
	topm  = 2 + StyleSizeToInt(ButtonControl.style.top)  + StyleSizeToInt(ButtonControl.style.height) + StyleSizeToInt(Form1.style.top)
	Set MyMenuControl = ShowMenu(leftm, topm, 140, "HandleMenuControl", MenuControlItems)
End Sub


' Handlerul executat la selectarea unei optiuni din meniul Page
Sub HandleMenuPage(html)
	If html="<HR>" Or Len(html)=0 Then Exit Sub
	MyMenuPage.Hide
	Set MyMenuPage = Nothing

	Select Case html     
		Case MenuPageItems(0) Call AddPage()
		Case MenuPageItems(1) Call DelPage()
		Case MenuPageItems(3) Call RenPage()
	End Select 
End Sub


' Handlerul executat la selectarea unei optiuni din meniul Page
Sub HandleMenuControl(html)
	If html="<HR>" Or Len(html)=0 Then Exit Sub
	MyMenuControl.Hide
	Set MyMenuControl = Nothing

	Select Case html     
		Case MenuControlItems(0) Call AddControl()
		Case MenuControlItems(1) Call DelControl()
		Case MenuControlItems(3) Call RenControl()
	End Select 
End Sub


' Activeaza/Dezactiveaza controalele de pagina
' pentru momentele in care se face Saving...
Sub ActivateControls(bActivate)
	ButtonPage.disabled = not bActivate
	ButtonControl.disabled = not bActivate
	If bActivate Then
		LangTexts.style.visibility = "inherit"
		LangButtons.style.visibility = "inherit"
		LangWait.style.visibility = "hidden"
	Else
		LangTexts.style.visibility = "hidden"
		LangButtons.style.visibility = "hidden"
		LangWait.style.visibility = "inherit"
	End If
End Sub


' Este apelata automat de raspunsul intors de server
Sub HandleReturnFromServer(strFromAction, bErrorOnServer, ExtraInfo)
	Call ActivateControls(true)
	If bErrorOnServer Then Exit Sub
	
	Select Case strFromAction
		Case "addpage"
			lngPageToPosition = ExtraInfo
			lngControlToPosition = -1
			TDCPage.DataURL = TDCPage.DataURL
			TDCPage.Reset
		Case "renpage"
			lngPageToPosition = ExtraInfo
			lngControlToPosition = -1
			TDCPage.DataURL = TDCPage.DataURL
			TDCPage.Reset
		Case "delpage"
			lngPageToPosition = 0
			lngControlToPosition = -1
			TDCPage.DataURL = TDCPage.DataURL
			TDCPage.Reset
		Case "addcontrol"
			lngPageToPosition = -1
			lngControlToPosition = ExtraInfo
			TDCControl.DataURL = TDCControl.DataURL
			TDCControl.Reset
		Case "rencontrol"
			lngPageToPosition = -1
			lngControlToPosition = ExtraInfo
			TDCControl.DataURL = TDCControl.DataURL
			TDCControl.Reset
		Case "delcontrol"
			lngPageToPosition = -1
			lngControlToPosition = 0
			TDCControl.DataURL = TDCControl.DataURL
			TDCControl.Reset
		Case "savetexts"
			TDCTexts.DataURL = TDCTexts.DataURL
			TDCTexts.Reset
			Call DisableControls(False)
	End Select
End Sub


' Adauga o noua pagina
Sub AddPage
	Dim strPageName
	
	strPageName = ShowModalDialog("inputbox_cli.asp?pagetitle=Add new form&label=Name:", , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
	If Not IsEmpty(strPageName) Then
		FormularH.SelectAction.value = "addpage"
		FormularH.SelectValue.value  =  strPageName
		Call ActivateControls(false)
		FormularH.submit
	End If
End Sub


' Sterge o pagina
Sub DelPage
	Dim PageID
	Dim strQuestion1
	Dim strQuestion2

	strQuestion1 = "Are you sure you want to delete selected form?"
	strQuestion2 = "Selected form has associated controls." & vbCrLf & vbCrLf & "Are you sure you want to delete selected form and all its controls?"
	If ComboBoxPage.selectedIndex > -1 Then
		If MsgBox(strQuestion1, vbYesNo + vbQuestion, "Warning") = vbNo Then Exit Sub
		If ListBoxControl.options.Length > 0 Then
			If MsgBox(strQuestion2, vbYesNo + vbQuestion, "Warning") = vbNo Then Exit Sub
		End If
		PageID = ComboBoxPage.value
		FormularH.SelectAction.value = "delpage"
		FormularH.SelectValue.value  =  PageID
		Call ActivateControls(false)
		FormularH.submit
	End If
End Sub


' Redenumeste o pagina
Sub RenPage()
	Dim PageID
	Dim PageName
	Dim NewPageName
	
	If ComboBoxPage.selectedIndex > -1 Then
		PageID		= ComboBoxPage.value
		PageName	= ComboBoxPage.options(ComboBoxPage.selectedIndex).Text

		NewPageName = ShowModalDialog("inputbox_cli.asp?pagetitle=Rename form&label=New name:&text=" & PageName, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
		If Not IsEmpty(NewPageName) Then
			If (Len(Trim(NewPageName)) > 0) and (NewPageName <> PageName) Then 
				FormularH.SelectAction.value = "renpage"
				FormularH.SelectList.value = PageID
				FormularH.SelectValue.value  =  NewPageName
				Call ActivateControls(false)
				FormularH.submit
			End If
		End If
	End If
End Sub


' Adauga un nou control
Sub AddControl
	Dim strControlName
	
	strControlName = ShowModalDialog("inputbox_cli.asp?pagetitle=Add control&label=Name:", , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
	If (Not IsEmpty(strControlName)) and ((ComboBoxPage.selectedIndex > -1)) Then
		FormularH.SelectAction.value = "addcontrol"
		FormularH.SelectValue.value  =  strControlName
		FormularH.SelectList.value  =  ComboBoxPage.value
		Call ActivateControls(false)
		FormularH.submit
	End If
End Sub


' Sterge un control
Sub DelControl
	Dim ControlID
	Dim strQuestion

	strQuestion = "Are you sure you want to delete selected control?"
	If ListBoxControl.selectedIndex > -1 Then
		If MsgBox(strQuestion, vbYesNo + vbQuestion, "Atentie") = vbNo Then Exit Sub
		ControlID = ListBoxControl.value
		FormularH.SelectAction.value = "delcontrol"
		FormularH.SelectValue.value  =  ControlID
		Call ActivateControls(false)
		FormularH.submit
	End If
End Sub



' Redenumeste un control
Sub RenControl()
	Dim ControlID
	Dim ControlName
	Dim NewControlName
	
	If ListBoxControl.selectedIndex > -1 Then
		ControlID		= ListBoxControl.value
		ControlName		= ListBoxControl.options(ListBoxControl.selectedIndex).Text

		NewControlName = ShowModalDialog("inputbox_cli.asp?pagetitle=Rename control&label=New name:&text=" & ControlName, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
		If Not IsEmpty(NewControlName) Then
			If (Len(Trim(NewControlName)) > 0) and (NewControlName <> ControlName) Then 
				FormularH.SelectAction.value = "rencontrol"
				FormularH.SelectList.value = ControlID
				FormularH.SelectValue.value  =  NewControlName
				Call ActivateControls(false)
				FormularH.submit
			End If
		End If
	End If
End Sub



' Trateaza evenimentul care apare la apasarea butonului Close
Sub ButtonClose_OnClick
	HideAllDivs
End Sub

' Trateaza evenimentul care apare la apasarea butonului Editeaza (texte)
Sub ButtonEdit_OnClick
	Dim id
	
	If lngIDControl < 0 Then
		MsgBox "You need to select a control first.", vbOkOnly + vbExclamation, "Info"
		Exit Sub
	End If
	
	For Each id In arLanguages
		document.all("txt_" & id).Disabled = False
	Next

	ButtonEditOK.disabled = False
	ButtonEditCancel.disabled = False
	ButtonEdit.Disabled = True
	Call DisableControls(True)
End Sub

' Trateaza evenimentul care apare la apasarea butonului Anuleaza editare (texte)
Sub ButtonEditCancel_OnClick
	Call FillTextsFromTDC()
	
	ButtonEditOK.Disabled = True
	ButtonEditCancel.Disabled = True
	ButtonEdit.Disabled = False
	Call DisableControls(False)
End Sub

' Trateaza evenimentul care apare la apasarea butonului Accepta editare (texte)
Sub ButtonEditOK_OnClick
	Dim texts
	Dim rs
	Dim s
	
	If ListBoxControl.selectedIndex = -1 Then
		MsgBox "You need to select a control first.", vbOkOnly + vbExclamation
		Exit Sub
	End If
	
	texts = ""
	For Each id In arLanguages
		s = id & Chr(1) & document.all("txt_" & id).Value
		texts = texts & s & Chr(2)
	Next

	If Len(texts) > 0 Then 
		texts = Left(texts, Len(texts) - Len(Chr(2)))

		FormularH.SelectAction.value = "savetexts"
		FormularH.SelectList.value = ListBoxControl.value
		FormularH.SelectValue.value  =  texts
		Call ActivateControls(false)
		FormularH.submit
	End If
End Sub


' Dezactiveaza/activeaza o parte din controalele de pe ecran
' pentru a nu interfera cu procesul de editare a textelor
Sub DisableControls(bDisable)
	ComboBoxPage.disabled = bDisable
	ListBoxControl.disabled = bDisable
	ButtonControl.disabled = bDisable
	ButtonPage.disabled = bDisable
End Sub
</script>

</body>
</html>

