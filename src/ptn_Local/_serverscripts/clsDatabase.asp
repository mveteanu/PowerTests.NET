<%
Class clsDatabase
	Private objCon

	' Gets or Sets the ADO connection
	' If setted to Nothing then the connection will be created internally
	Public Property Get Connection
		Call BuildInternal()
		Set Connection = objCon
	End Property
	Public Property Set Connection(ByRef cn)
		Set objCon = cn
	End Property

	' Transactional support
	Public Sub BeginTrans
		Call BuildInternal()
		objCon.BeginTrans()
	End Sub
	Public Sub CommitTrans
		Call BuildInternal()
		objCon.CommitTrans()
	End Sub
	Public Sub RollbackTrans
		Call BuildInternal()
		objCon.RollbackTrans()
	End Sub
	
	
	' If the stored proc doesn't take parameters
	' varParams should be Empty
	Public Function GetRS(strStoredProc, varParams)
		Dim cmd
		Dim rs
		
		Call BuildInternal()
		Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection = objCon
		cmd.CommandText = strStoredProc
		If IsEmpty(varParams) Then
			Set rs = cmd.Execute
		Else
			Set rs = cmd.Execute(,varParams,adCmdStoredProc)
		End If
		
		Set cmd = Nothing
		Set GetRS = rs
	End Function
	
	Public Function GetScalar(strField, strStoredProc, varParams)
		Dim rs
		
		Set rs = GetRS(strStoredProc, varParams)
		If Not rs.EOF Then GetScalar = rs.Fields(strField).Value
	End Function


	' Sends to DB a SQL/named stored proc for execution
	Public Sub ExecCommand(strSQL, varParams)
		Dim cmd
		
		Call BuildInternal()		
		Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection = objCon
		cmd.CommandText = strSQL

		If IsEmpty(varParams) Then
			Call cmd.Execute
		Else
			Call cmd.Execute(,varParams,adCmdStoredProc)
		End If

		Set cmd = Nothing
	End Sub


	' Returns the ID of the last updated record (or 0 if cannot be obtained)
	Public Function GetLastID(strTableName)
		Call BuildInternal()
		GetLastID = CLng((objCon.Execute("SELECT @@IDENTITY FROM " & strTableName)).Fields(0).Value)
	End Function

	' Starts private area
	
	Private Sub BuildInternal()
		If objCon Is Nothing Then
			Set objCon = Server.CreateObject("ADODB.Connection")
			objCon.Open Application("DSN")
		End If	
	End Sub
	
	Private Sub Class_Initialize()
		Set objCon = Nothing
	End Sub

	Private Sub Class_Terminate()
		If Not(objCon is Nothing) Then
			If objCon.State = adStateOpened Then objCon.Close
			Set objCon = Nothing
		End If
	End Sub
End Class
%>