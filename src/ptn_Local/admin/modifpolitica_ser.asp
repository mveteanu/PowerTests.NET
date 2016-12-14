<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/settings.asp" -->
<%
Dim objCon, ReturnedError
ReturnedError = ""

On Error Resume Next

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.CursorLocation = adUseClient
objCon.Open Application("DSN")
objCon.BeginTrans 

Call SetUserPolicy("A", Request.Form("adminpolicy").Item, objCon)
Call SetUserPolicy("P", Request.Form("profpolicy").Item,  objCon)
Call SetUserPolicy("S", Request.Form("studpolicy").Item,  objCon)

If Err.number <> 0 then
	objCon.RollbackTrans 
	ReturnedError = """Error (" & Err.number & "): " & Err.Description & """, vbOkOnly + vbCritical"
	Err.Clear 
Else
	objCon.CommitTrans 
End If
 
objCon.Close
Set objCon = Nothing

Call PrintClientResponse

' Trimite spre client un raspuns ca rezultat al salvarii datelor
Sub PrintClientResponse
	Dim strMsg
	strMsg = "Data was saved."
	With Response
		.Write "<script language=vbscript>" & vbCrLf
		.Write "window.parent.ActivateButtons true" & vbCrLf
		If Len(ReturnedError) = 0 Then
			.Write "msgbox """ & strMsg & """, vbOkOnly + vbInformation" & vbCrLf
		Else
			.Write "msgbox " & ReturnedError & vbCrLf
		End If
		.Write "</script>" & vbCrLf
	End With
End Sub 
%>
