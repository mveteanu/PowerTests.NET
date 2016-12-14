<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/settings.asp" -->
<%
Const MsgAddPageOK	= """Form was added succesfully."", vbOkOnly+vbInformation"
Const MsgRenPageOK	= """Form was renamed succesfully."", vbOkOnly+vbInformation"
Const MsgDelPageOK	= """Form was deleted."", vbOkOnly+vbInformation"
Const MsgAddControlOK	= """Control was added succesfully."", vbOkOnly+vbInformation"
Const MsgRenControlOK	= """Control was renamed succesfully."", vbOkOnly+vbInformation"
Const MsgDelControlOK	= """Control was deleted."", vbOkOnly+vbInformation"
Const MsgSaveTextsOK	= """Texts were saved."", vbOkOnly+vbInformation"
Const MsgError		= """DB errors occured."", vbOkOnly+vbCritical"

Dim action
Dim cn, db
Dim Mesaj
Dim ExtraInfo

On Error Resume next

ExtraInfo = "0"
action = LCase(Request.Form("SelectAction").Item)
Set cn = server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
Set db = New clsDatabase
Set db.Connection = cn

Select Case action
	Case "addpage"
		Call AddPage(Request.Form("SelectValue").Item)
		Mesaj = MsgAddPageOK
	Case "delpage"
		Call DelPage(Request.Form("SelectValue").Item)
		Mesaj = MsgDelPageOK
	Case "renpage"
		Call RenPage(Request.Form("SelectList").Item, Request.Form("SelectValue").Item)
		Mesaj = MsgRenPageOK
	Case "addcontrol"
		Call AddControl(Request.Form("SelectValue").Item, Request.Form("SelectList").Item)
		Mesaj = MsgAddControlOK
	Case "delcontrol"
		Call DelControl(Request.Form("SelectValue").Item)
		Mesaj = MsgDelControlOK
	Case "rencontrol"
		Call RenControl(Request.Form("SelectList").Item, Request.Form("SelectValue").Item)
		Mesaj = MsgRenControlOK
	Case "savetexts"
		Call SaveTexts(Request.Form("SelectList").Item, Request.Form("SelectValue").Item)
		Mesaj = MsgSaveTextsOK
End Select

Set db = Nothing
cn.Close
Set cn = Nothing
 
If Err.number <> 0 Then
	Mesaj = MsgError
	Err.Clear 
End If

' ======================================================================

' Adauga o noua definitie de pagina de texte
Sub AddPage(strName)
	Dim sql
	
	sql = "INSERT INTO TBTextsPages(name) VALUES('" & strName & "')"
	Call db.ExecCommand(sql, Empty)
	ExtraInfo = db.GetLastID("TBTextsPages")
End Sub


' Sterge o pagina
Sub DelPage(lngPageID)
	Dim sql
	
	sql = "DELETE FROM TBTextsPages WHERE id = " & lngPageID
	Call db.ExecCommand(sql, Empty)
End Sub


' Redenumeste o pagina
Sub RenPage(lngPageID, strNewName)
	Dim sql
	
	sql = "UPDATE TBTextsPages SET name = '" & strNewName & "' WHERE id = " & lngPageID
	Call db.ExecCommand(sql, Empty)
	ExtraInfo = lngPageID
End Sub

' ======================================================================

' Adauga o noua definitie de control
Sub AddControl(strName, lngPageID)
	Dim sql
	Dim rs
	
	On Error Resume Next
	Call db.BeginTrans()
	
	sql = "INSERT INTO TBTextsControls(controlname, id_page) VALUES('" & strName & "', " & lngPageID & ")"
	Call db.ExecCommand(sql, Empty)
	ExtraInfo = db.GetLastID("TBTextsControls")
	
	Set rs = db.GetRS("SELECT id FROM TBLanguages", Empty)
	Do While Not rs.EOF
		sql = "INSERT INTO TBTexts(id_control, id_lang, textvalue) VALUES(" & ExtraInfo & ", @1, '-')"
		Call db.ExecCommand(Replace(sql, "@1", rs.Fields("id").Value), Empty)
		rs.MoveNext
	Loop
	Set rs = Nothing

	If Err.number <> 0 Then
		Call db.RollbackTrans() 
	Else
		Call db.CommitTrans()
	End If
End Sub


' Sterge un control si textele asociate
Sub DelControl(lngControlID)
	Dim sql
	
	sql = "DELETE FROM TBTextsControls WHERE id = " & lngControlID
	Call db.ExecCommand(sql, Empty)
End Sub


' Redenumeste un control
Sub RenControl(lngControlID, strNewName)
	Dim sql
	
	sql = "UPDATE TBTextsControls SET controlname = '" & strNewName & "' WHERE id = " & lngControlID
	Call db.ExecCommand(sql, Empty)
	ExtraInfo = lngControlID
End Sub

' ======================================================================

' Salveaza textele
Sub SaveTexts(lngControlID, strPackedValues)
	Dim arAllTexts
	Dim arText
	Dim s
	Dim sql
	
	arAllTexts = Split(strPackedValues, Chr(2))
	For Each s In arAllTexts
		arText = Split(s, Chr(1))

		sql = "UPDATE TBTexts SET textvalue = '" & arText(1) & "' WHERE id_control = " & lngControlID & " AND id_lang = " & arText(0)
		Call db.ExecCommand(sql, Empty)
	Next
End Sub

%>
<body>
<script language=vbscript>
	Call window.parent.HandleReturnFromServer("<%=action%>", <%=CBool(Mesaj = MsgError)%>, <%=ExtraInfo%>)
	MsgBox <%=Mesaj%>
</script>
</body>
