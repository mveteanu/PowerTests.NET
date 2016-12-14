<%@ Language=VBScript %>
<%
Dim ReturnedError, cn
ReturnedError = ""

On Error Resume Next
 
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
cn.BeginTrans 
 
Select Case LCase(Request.Form("SelectAction"))
	Case "reg"		
		RegMessages = Array("""Course enrollment was completed successfully."", vbOKOnly+vbInformation", _
							"""Course enrollment request was received successfully."" & vbCrLf & ""Please wait for professor to validate your request."", vbOKOnly+vbInformation", _
							"""You cannot enroll to selected course because"" & vbCrLf & ""the course has reached the maximum students."", vbOKOnly+vbExclamation", _
							"""You cannot enroll to selected course because"" & vbCrLf & ""professor blocked new enrollemnts."", vbOKOnly+vbExclamation", _
							"""You cannot enroll to selected course because"" & vbCrLf & ""you are allready enrolled in this class."", vbOKOnly+vbExclamation" _
							)
		Mesaj = RegMessages(GetRegistrationStatus(Request.Form("SelectList"), Session("UserID"), cn)) 
	Case "unsubscr"	
		UnsubscribeFromCurs Request.Form("SelectList"), Session("UserID"), cn
		Mesaj = """Un-enrollment completed successfully."", vbOKOnly+vbInformation"
End Select

If Err.number <> 0 then
	cn.RollbackTrans 
	ReturnedError = """Error (" & Err.number & "): " & Err.Description & """, vbOkOnly + vbCritical"
	Err.Clear 
Else
	cn.CommitTrans 
End If

cn.Close
Set cn = nothing

Call PrintClientResponse

' Incearca sa inscrie un student la un curs si intoarce
' starea operatiei efectuate in functie de care se alege
' mesajul returnat utilizatorului
Function GetRegistrationStatus(cursid, studid, objCon)
	Dim re, rs1, rs2, myCmd
  
	If GetNrInscrieri(cursid, studid, objCon) = 0 Then
		Set myCmd = Server.CreateObject("ADODB.Command")
		Set myCmd.ActiveConnection = objCon
		myCmd.CommandText = "GetCursByID"
		myCmd.CommandType = adCmdStoredProc
		Set rs2 = myCmd.Execute(,CLng(cursid))
		Set myCmd = Nothing
		Select Case rs2.Fields("permisiiacceptare").Value 
			Case 0 
				If (rs2.Fields("StudentiInscrisi").Value >= rs2.Fields("maxstudents").Value) and (rs2.Fields("maxstudents").Value<>0) then
					re = 2
				Else
					objCon.StudentiLaCursuri_SubscribeReg CLng(studid), CLng(cursid)
					re = 0
				End If
			Case 1 
				objCon.StudentiLaCursuri_SubscribeOnly CLng(studid), CLng(cursid)
				re = 1
			Case 2 
				re = 3
		End Select
		rs2.Close
		Set rs2 = Nothing
	Else
		re = 4  
	End If
	GetRegistrationStatus = re
End Function

' Intoarce nr. de inscrieri ale unui student la un curs
' In mod normal trebuie sa fie 0 sau 1
Function GetNrInscrieri(cursid, studid, objCon)
	Dim rs, re
	Set rs = Server.CreateObject("ADODB.Recordset")
	objCon.StudentiLaCursuri_NrInscrieri CLng(studid), CLng(cursid), rs
	re = rs.Fields("NrInscrieri").Value 
	rs.Close
	Set rs = Nothing
	GetNrInscrieri = re
End Function

' Desubscrie un student de la un anumit curs
Sub UnsubscribeFromCurs(cursid, studid, objCon)
	objCon.StudentiLaCursuri_Unsubscribe CLng(studid), CLng(cursid)
End Sub

' Trimite spre client un raspuns ca rezultat al salvarii datelor
Sub PrintClientResponse
	With Response
		.Write "<script language=vbscript>" & vbCrLf
		If Len(ReturnedError) = 0 Then
			.Write "window.parent.ReloadTDC"  & vbCrLf
			.Write "msgbox " & Mesaj & vbCrLf
		Else
			.Write "window.parent.ActivateButtons true" & vbCrLf
			.Write "msgbox " & ReturnedError & vbCrLf
		End If
		.Write "</script>" & vbCrLf
	End With
End Sub 
%>
