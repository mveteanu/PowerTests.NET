<%@ Language=VBScript %>
<%
Dim objCon, ReturnedError, Action

On Error Resume Next

ReturnedError = ""
Action		  = Request.Form("SelectAction")
CereriLst	  = Request.Form("SelectList")

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.CursorLocation = adUseClient
objCon.Open Application("DSN")
objCon.BeginTrans 

Select Case Action
	Case "accept"	Call AcceptaCereriStudenti(CereriLst, objCon)
	Case "deny"		Call RespingeCereriStudenti(CereriLst, objCon)
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

' Accepta cererile unei liste de studenti de inscriere la cursuri.
Sub AcceptaCereriStudenti(idcerere, objCon)
	Dim ArCereri, Cerere
	ArCereri = Split(idcerere, ",", -1, 1)
	For Each Cerere In ArCereri
		objCon.CereriStudenti_AcceptStud CLng(Cerere)
	Next
End Sub

' Respinge cererile de inscriere la cursuri ale unei liste de studenti.
Sub RespingeCereriStudenti(idcerere, objCon)
	Dim ArCereri, Cerere
	ArCereri = Split(idcerere, ",", -1, 1)
	For Each Cerere In ArCereri
		objCon.CereriStudenti_DenyStud CLng(Cerere)
	Next
End Sub

' Trimite spre client un raspuns ca rezultat al salvarii datelor
Sub PrintClientResponse
	With Response
		.Write "<script language=vbscript>" & vbCrLf
		If Len(ReturnedError) = 0 Then
			Select Case Action
				Case "accept"	strMsg = "Selected students have received permission to attend the course."
				Case "deny"		strMsg = "Selected students were denied to attend this course."
			End Select
			.Write "window.parent.ReloadTDC"  & vbCrLf
			.Write "msgbox """ & strMsg & """, vbOkOnly + vbInformation" & vbCrLf
		Else
			.Write "window.parent.ActivateButtons true, true, true, true, true" & vbCrLf
			.Write "msgbox " & ReturnedError & vbCrLf
		End If
		.Write "</script>" & vbCrLf
	End With
End Sub 
%>
