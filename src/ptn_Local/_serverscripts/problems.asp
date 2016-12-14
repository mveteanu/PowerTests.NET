<%
' Returns a Dictionary object with the questions from a certain category
' DictKey = QuestionId
' DictItem = QuestionName
Function GetCategPBList(categid, objCon)
 Dim re
 
 Set re = CreateObject("Scripting.Dictionary")
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetPbFromCategID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(categid))
 If not rs.EOF then
   do while not rs.EOF
     re.Add CStr(rs.Fields("id_problem").Value), rs.Fields("numeproblema").Value
     rs.MoveNext
   loop
 End If  
 rs.Close
 set rs = nothing
 set myCmd = nothing
 
 Set GetCategPBList = re
End Function


' Returns a Dictionary object with the questions from a certain course
' DictKey = QuestionId
' DictItem = QuestionName
Function GetTotalPBList(cursid, objCon)
 Dim re
 
 Set re = CreateObject("Scripting.Dictionary")
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetPbInfo"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(cursid))
 If not rs.EOF then
   do while not rs.EOF
     re.Add CStr(rs.Fields("id_problem").Value), rs.Fields("numeproblema").Value
     rs.MoveNext
   loop
 End If  
 rs.Close
 set rs = nothing
 set myCmd = nothing
 
 Set GetTotalPBList = re
End Function


' Returns a Dictionary object with categories from a certain course
Function GetCategDiction(cursid, objCon)
 Dim re
 
 Set re = CreateObject("Scripting.Dictionary")
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetPbCategByCursID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(cursid))
 do until rs.EOF
   re.Add CStr(rs.Fields("id_categpb").value), rs.Fields("numecateg").value
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 Set GetCategDiction = re
End Function


' Returns an array with all IDs of questions of a course
Function GetTotalPbArray(cursid, objCon)
  Dim re
  
  Redim re(-1)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbInfo"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(cursid))
  Do Until rs.EOF
   re = AddItemToArray(re, rs.Fields("id_problem").Value)
   rs.MoveNext
  Loop
  rs.Close
  Set rs = nothing
  Set myCmd = nothing
  
  GetTotalPbArray = re
End Function


' Returns an array with IDs of all questions from a course
Function GetCategPbArray(categid, objCon)
  Dim re
  
  Redim re(-1)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbFromCategID"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(categid))
  Do Until rs.EOF
   re = AddItemToArray(re, rs.Fields("id_problem").Value)
   rs.MoveNext
  Loop
  rs.Close
  Set rs = nothing
  Set myCmd = nothing
  
  GetCategPbArray = re
End Function
%>
