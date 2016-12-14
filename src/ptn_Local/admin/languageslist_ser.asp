<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/settings.asp" -->
<%
Const MesajTest = """OK"""
Const MsgAddOK = """The new language was added succesfully."", vbOkOnly+vbInformation"
Const MsgRenOK = """The language was renamed succesfully."", vbOkOnly+vbInformation"
Const MsgDelOK = """The language and corresponding text were deleted."", vbOkOnly+vbInformation"
Const MsgPrefOK = """System default language was set."", vbOkOnly+vbInformation"
Const MsgError = """DB errors occured."", vbOkOnly+vbCritical"

Dim cn, db
Dim Mesaj

On Error Resume next
 
Set cn = server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
Set db = New clsDatabase
Set db.Connection = cn

Select Case LCase(Request.Form("SelectAction").Item)
	Case "add"
		Call AddLanguage(Request.Form("SelectValue").Item, Request.Form("SelectList").Item)
		Mesaj = MsgAddOK
	Case "ren"
		Call RenLanguage(Request.Form("SelectList").Item, Request.Form("SelectValue").Item)
		Mesaj = MsgRenOK
	Case "del"
		Call DelLanguage(Request.Form("SelectList").Item, Request.Form("SelectValue").Item)
		Mesaj = MsgDelOK
	Case "preflang"
		Call SetMainLanguage(Request.Form("SelectList").Item)
		Mesaj = MsgPrefOK
	Case Else
		Mesaj = MesajTest
End Select

Set db = Nothing
cn.Close
Set cn = Nothing
 
If Err.number <> 0 then
  Mesaj = MsgError
  Err.Clear 
end if

' Adauga o noua limba in baza de date. Se specifica numele noii limbi 
' si ID-ul unei alte limbi in cazul in care se doreste copierea textelor
' din acea limba in noua limba ce se creaza
Sub AddLanguage(strName, lngLangBase)
	Dim sql
	Dim txtValue
	Dim lngIDBase
	Dim lngIDNewLang
	
	On Error Resume Next
	Call db.BeginTrans()
	sql = "INSERT INTO TBLanguages(langname) VALUES('" & strName & "')"
	Call db.ExecCommand(sql, Empty)
	lngIDNewLang = db.GetLastID("TBLanguages")
	
	If lngLangBase > 0 Then
		txtValue  = "textvalue"
		lngIDBase = lngLangBase
	Else
		txtValue  = "'-'"
		lngIDBase = GetPreferredLanguage(db.Connection)
	End If
	
	sql = "INSERT INTO TBTexts (id_control, id_lang, textvalue) " &_
	"SELECT id_control, " & lngIDNewLang & ", " & txtValue & " FROM TBTexts WHERE id_lang = " & CStr(lngIDBase)
	Call db.ExecCommand(sql, Empty)
	
	If Err.number <> 0 Then
		Call db.RollbackTrans() 
	Else
		Call db.CommitTrans()
	End If
	
End Sub

' Rename a language
Sub RenLanguage(idLanguage, strNewName)
	Dim sql
	
	sql = "UPDATE TBLanguages SET langname = '" & Left(strNewName,50) & "' WHERE id = " & CStr(idLanguage)
	Call db.ExecCommand(sql, Empty)
End Sub


' Sterge o limba impreuna cu textele asociate
' Daca limba este folosita de utilizatori.. atunci acestora li
' se va atribui noua limba dupa stergere
Sub DelLanguage(idLanguage, idNewLang)
	Dim sql
	
	On Error Resume Next
	Call db.BeginTrans()

	sql = "DELETE FROM TBLanguages WHERE id = " & idLanguage
	Call db.ExecCommand(sql, Empty)
	If CLng(idNewLang) > 0 Then
		sql = "UPDATE TBUsers SET id_lang = " & idNewLang & " WHERE id_lang = " & idLanguage
		Call db.ExecCommand(sql, Empty)
	End If

	If Err.number <> 0 Then
		Call db.RollbackTrans() 
	Else
		Call db.CommitTrans()
	End If
End Sub

' Seteaza limba principala din aplicatie
Sub SetMainLanguage(idLang)
	Call SetPreferredLanguage(idLang, db.Connection)
End Sub
%>
<body>
<script language=vbscript>
	window.parent.ReloadTDC
	msgbox <%=Mesaj%>
</script>
</body>
