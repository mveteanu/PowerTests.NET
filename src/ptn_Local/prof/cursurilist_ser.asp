<%@ Language=VBScript %>
<%
Dim objCon, ReturnedError, Action

On Error Resume Next

ReturnedError = ""
Action      = Request.Form("SelectAction")
CursuriList = Request.Form("SelectList")
CursVals    = Split(Request.Form("SelectValues"),"|",-1,1)

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.CursorLocation = adUseClient
objCon.Open Application("DSN")
objCon.BeginTrans 

Select Case Action
	Case "del"	Call DeleteCursuri(CursuriList, objCon)
	Case "add"	Call AddCurs(Session("UserID"), CursVals, objCon)
	Case "edit"	Call UpdateCurs(CursuriList,CursVals, objCon)
End Select

If Err.number <> 0 then
	objCon.RollbackTrans 
	ReturnedError = """Error (" & Err.number & "): " & Err.Description & """, vbOkOnly + vbCritical"
	Err.Clear 
Else
	objCon.CommitTrans 
End If
 
objCon.Close
Set objCon = nothing

Call PrintClientResponse()

' Adauga un nou curs cu datele specificate prin CursData
' CursData este un array ce contine valorile
'   CursData(0) = numecurs
'   CursData(1) = maxstudents
'   CursData(2) = permisiiacceptare
'   CursData(3) = public
Sub AddCurs(userID, CursData, objCon)
	objCon.CursuriAdd CLng(userID), CStr(CursData(0)), Abs(CLng(CursData(1))), CByte(CursData(2)), CBool(CursData(3))
End Sub

' Modifica datele cursului cu ID-ul dat prin cursID
' CursData este un array ce contine noile valori
'   CursData(0) = numecurs
'   CursData(1) = maxstudents
'   CursData(2) = permisiiacceptare
'   CursData(3) = public
Sub UpdateCurs(cursID, CursData, objCon)
	objCon.CursuriEdit CLng(cursID), CStr(CursData(0)), Abs(CLng(CursData(1))), CByte(CursData(2)), CBool(CursData(3))
End Sub

' Sterge inregistrarile cursurilor specificate prin cursList
Sub DeleteCursuri(cursList, objCon)
	Const SQLDelete = "DELETE FROM TBCursuri WHERE id_curs IN (@1)"
	objCon.Execute Replace(SQLDelete, "@1", cursList)
End Sub

' Trimite spre client un raspuns ca rezultat al salvarii datelor
Sub PrintClientResponse
	With Response
		.Write "<script language=vbscript>" & vbCrLf
		If Len(ReturnedError) = 0 Then
			Select Case Action
				Case "del"	strMsg = "Selected course and associated date were successfully deleted."
				Case "add"	strMsg = "Course created successfully."
				Case "edit"	strMsg = "Course properties were successfully saved."
			End Select
			.Write "window.parent.ReloadTDC"  & vbCrLf
			.Write "msgbox """ & strMsg & """, vbOkOnly + vbInformation" & vbCrLf
		Else
			.Write "window.parent.ActivateButtons true" & vbCrLf
			.Write "msgbox " & ReturnedError & vbCrLf
		End If
		.Write "</script>" & vbCrLf
	End With
End Sub 
%>
