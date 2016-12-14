<%
' Returns the value of a variable stored in DB
Function GetApplicationSettings(settingsvariable,cn)
	Dim re
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = cn
	cmd.CommandText = "GetApplicationSettings"
	Set rs = cmd.Execute(,settingsvariable,adCmdStoredProc)
	If not rs.EOF Then re = rs.Fields("valoarevariabila").Value
	rs.Close
	Set cmd = Nothing
	Set rs = Nothing
	
	GetApplicationSettings = re
End Function

' Sets the value of a DB storred variable
Sub SetApplicationSettings(numevariabila,valoarevariabila,cn)
	Const SQL_UPDATESET = "UPDATE TBApplicationSettings SET valoarevariabila='@1' WHERE numevariabila='@2'"
	cn.Execute Replace(Replace(SQL_UPDATESET,"@2",CStr(numevariabila)),"@1",CStr(valoarevariabila))
End Sub


' Sets the preffered application language
Function SetPreferredLanguage(idLang, objCon)
	Call SetApplicationSettings("preferredlanguage", idLang, objCon)
End Function


' Retursn application preferred language
Function GetPreferredLanguage(objCon)
	Dim re
	Dim cn
	Dim bIsIntern
	
	If objCon is Nothing Then
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.Open Application("DSN")
		bIsIntern = true
	Else
		Set cn = objCon
		bIsIntern = false
	End If
	
	re = CLng(GetApplicationSettings("preferredlanguage", cn))
	If bIsIntern Then
		cn.Close
		Set cn = Nothing
	End If
	 
	 GetPreferredLanguage = re
End Function


' Returns user validation policy based on user type: "A", "P", "S"
Function GetUserPolicy(tipuser,cn)
	Dim settingsvariable

	Select case tipuser
		Case "A" settingsvariable = "acceptadminpolicy"
		Case "P" settingsvariable = "acceptprofpolicy"
		Case "S" settingsvariable = "acceptstudpolicy"
	End Select

	GetUserPolicy = CInt(GetApplicationSettings(settingsvariable,cn))
End Function


' Sets user validation policy for a certain user type: "A", "P", "S"
' Returns true for success
Sub SetUserPolicy(tipuser,policy,cn)
	Dim settingsvariable

	Select case tipuser
		Case "A" settingsvariable = "acceptadminpolicy"
		Case "P" settingsvariable = "acceptprofpolicy"
		Case "S" settingsvariable = "acceptstudpolicy"
	End select

	Call SetApplicationSettings(settingsvariable, policy, cn)
End Sub
%>
