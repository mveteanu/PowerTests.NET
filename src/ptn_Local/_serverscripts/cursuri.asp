<%
' Returns an recordset with courses of a specified professor
Public Function GetCursuriByProf(id,objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetCursuriByProf"
  Set GetCursuriByProf = myCmd.Execute(,CLng(id))
  Set myCmd = Nothing
End Function


' Returns an Recordset with properties of specified course
Public Function GetCursByID(id,objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetCursByID"
  Set GetCursByID = myCmd.Execute(,CLng(id))
  Set myCmd = Nothing
End Function
%>
