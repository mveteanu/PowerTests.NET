<!-- #include file="utils.asp" -->
<!-- #include file="emails.asp" -->
<%
' Returns a RecordSet with user data based on his login info
Public Function GetUser(login,pass,objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetUser"
  Set GetUser = myCmd.Execute(,Array(CStr(login),CStr(pass)))
  Set myCmd = Nothing
End Function


' Returns a RecordSet with user data based on his Id
Public Function GetUserByID(id,objCon)
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetUserByID"
  Set GetUserByID = myCmd.Execute(,CLng(id))
  Set myCmd = Nothing
End Function


' Returns a Recordset with users data
Public Function GetUsersByIDs(idlist, objCon)
  Const SQLStr= "SELECT TBUsers.*, TBPersons.* FROM TBUsers INNER JOIN TBPersons ON TBUsers.id_person=TBPersons.id_person WHERE TBUsers.id_user IN (@1)"
  Set GetUsersByIDs = objCon.Execute(Replace(SQLStr, "@1", idlist))
End Function


' Validates a user (iduser) and sends him an email
' Returns
'          0 = OK
'          1 = DB validation error (no email was sent)
'          2 = DB validation OK but cannot send email
Function ValidateUser(iduser, idadmin, objCon)
  Const SQL_UPDATEUS = "UPDATE TBUsers SET datavalidare='@1' WHERE id_user=@2"
  Dim re
  re = 0

  On Error Resume Next
  objCon.Execute Replace(Replace(SQL_UPDATEUS, "@2", iduser), "@1", Now)
  If Err.number<>0 then
    Err.Clear 
    re = re or 1
  Else
    If not SendConfirmationEmail(iduser, idadmin, 1, objCon) then re = re or 2
  End If
  ValidateUser = re
End Function


' Deletes the request of a user
' Returns:
'          0 = OK
'          1 = Cannot send email
'          2 = Error deleting from DB (Attention: the email is sent first and then tries to delete !?...)
'          3 = Error sending email and deleting from DB
Function DeleteUnvalidatedUser(iduser, idadmin, objCon)
  Const SQL_DELUS = "DELETE FROM TBUsers WHERE id_user=@1"
  Dim re
  re = 0

  If not SendConfirmationEmail(iduser, idadmin, 2, objCon) then re = re or 1

  On Error Resume Next
  objCon.Execute Replace(SQL_DELUS,"@1",iduser)
  If Err.number<>0 then
    Err.Clear 
    re = re or 2
  End If
  DeleteUnvalidatedUser = re
End Function


' Functia valideaza toti utilizatori care sunt nevalidati si le 
' trimite emailuri de instiintare de la adresa administratorului
' Intoarce:
'          0 = OK
'          1 = Nu s-a putut obtine lista cu utilizatori ce necesita validare
'          2 = Unii utilizatori (sau toti) nu au putut fi validati in BD deci nu li s-a trimis nici emailuri
'          4 = La unii utilizatori (sau la toti) nu li s-a putut trimite email de instiintare
Function ValidateAllUsers(idadmin, objCon)
 Dim re
 re = 0
 
 On Error Resume Next
 Set cmd = Server.CreateObject("ADODB.Command")
 Set cmd.ActiveConnection = objCon
 cmd.CommandText = "GetUnvalidatedUsers"
 cmd.CommandType = adCmdStoredProc
 set rs=cmd.Execute
 If Err.number<>0 then
   re = re or 1
   Err.Clear
 Else
   do until rs.EOF
     select case ValidateUser(rs.Fields("id_user").value,idadmin,objCon)
       case 1 re = re or 2
       case 2 re = re or 4
     end select
     rs.movenext
   loop
 End If

 if rs.State = adStateOpen then rs.Close
 set rs=nothing
 set cmd=nothing
 ValidateAllUsers = re
End Function


' Functia sterge toti utilizatori care nu sunt nevalidati si le 
' trimite emailuri de instiintare de la adresa administratorului
' Intoarce:
'          0 = OK
'          1 = Nu s-a putut obtine lista cu utilizatori ce necesita validare
'          2 = Unii utilizatori (sau toti) nu au putut fi stersi (atentie: totusi au primit emailuri ca sunt stersi)
'          4 = La unii utilizatori (sau la toti) nu li s-a putut trimite email de instiintare
'          6 = Unii utiliz. nu au putut fi nici stersi nici nu li s-a trimis emailuri de instiintare
Function DeleteAllUnvalidatedUsers(idadmin, objCon)
 Dim re
 re = 0
 
 On Error Resume Next
 Set cmd = Server.CreateObject("ADODB.Command")
 Set cmd.ActiveConnection = objCon
 cmd.CommandText = "GetUnvalidatedUsers"
 cmd.CommandType = adCmdStoredProc
 set rs=cmd.Execute
 If Err.number<>0 then
   re = re or 1
   Err.Clear
 Else
   do until rs.EOF
     select case DeleteUnvalidatedUser(rs.Fields("id_user").value,idadmin,objCon)
       case 1 re = re or 4
       case 2 re = re or 2
       case 3 re = re or 6
     end select
     rs.movenext
   loop
 End If

 if rs.State = adStateOpen then rs.Close
 set rs=nothing
 set cmd=nothing
 DeleteAllUnvalidatedUsers = re
End Function


' Functia valideaza o lista de utilizatori si le trimite email-uri de instiintare
' Intoarce:
'          0 = OK
'          1 = Unii utilizatori (sau toti) nu au putut fi validati in BD (in acest caz nu se mai trimite nici email)
'          2 = Unii utilizatori (sau toti) au fost validati in BD dar nu li s-a putut trimite email
Function ValidateUsersList(usrlist, idadmin, objCon)
 Dim UsersList
 Dim idu
 Dim re

 UsersList = Split(usrlist,",",-1,1)
 re = 0
 for each idu in UsersList
   select case ValidateUser(idu, idadmin, objCon)
     case 1 re = re or 1
     case 2 re = re or 2
   end select
 next
End Function



' Functia sterge o lista de utilizatori nevalidati din BD si le trimite emailuri de instiintare
' Intoarce:
'          0 = OK
'          1 = La unii utilizatori (sau la toti) nu li s-a putut trimite emailuri de confirmare
'          2 = Unii utilizatori sau toti nu s-a sters din BD (Atentie: intai se trimite emailul si apoi se face stergerea?!.....)
'          3 = La unii utiliz. nu s-a putut trimite nici email si nici nu s-a reusit stergerea din BD
Function DeleteUnvalidatedUsersList(usrlist, idadmin, objCon)
 Dim UsersList
 Dim idu
 Dim re

 UsersList = Split(usrlist,",",-1,1)
 re = 0
 for each idu in UsersList
   select case DeleteUnvalidatedUser(idu, idadmin, objCon)
     case 1 re = re or 1
     case 2 re = re or 2
     case 3 re = re or 3
   end select
 next
End Function



' Functia intoarce un RecordSet ce contine lista cu utilizatori nevalidati in functie de tipul acestora
' tipUser = tip-ul utilizatorilor (A,P,S). Daca este null se intorc inregistrarile la toti userii nevalidati
Function GetUnvalidatedUsers(tipUser, objCon)
 set cmd=Server.CreateObject("ADODB.Command")
 Set cmd.ActiveConnection = objCon
 cmd.CommandText = "GetUnvalidatedUsers"
 cmd.CommandType = adCmdStoredProc
 set rs = cmd.Execute
 if tipUser<>"" then rs.Filter = "tipuser = '" & tipUser & "'"
 set GetUnvalidatedUsers = rs
 set cmd=nothing
End Function



' Functia intoarce un array cu 9 elemente avand urmatoarea semnificatie:
' element(0) = nr. de cereri pentru cont de administrator
' element(1) = nr. de cereri pentru cont de profesor
' element(2) = nr. de cereri pentru cont de student
'
' element(3) = nr. total de administratori inregistrati
' element(4) = nr. total de profesori inregistrati
' element(5) = nr. total de studenti inregistrati
'
' element(6) = nr. de administratori blocati
' element(7) = nr. de profesori blocati
' element(8) = nr. de studenti blocati
'
' In caz de eroare toate elementele contin valoarea -1
Function GetSumarUsers(objCon)
 Dim SumarUsers(9)
 
 on error resume next
 set cmd=Server.CreateObject("ADODB.Command")
 Set cmd.ActiveConnection = objCon
 cmd.CommandText = "GetNoOfUsers"
 cmd.CommandType = adCmdStoredProc
 set rs = cmd.Execute

 do until rs.EOF
  select case rs.Fields("tipuser").value
    case "A" SumarUsers(0) = NullToZero(rs.Fields("nevalidati").value)
             SumarUsers(3) = NullToZero(rs.Fields("users").value)
             SumarUsers(6) = NullToZero(rs.Fields("lockedusers").value)  
    case "P" SumarUsers(1) = NullToZero(rs.Fields("nevalidati").value)
             SumarUsers(4) = NullToZero(rs.Fields("users").value)
             SumarUsers(7) = NullToZero(rs.Fields("lockedusers").value)  
    case "S" SumarUsers(2) = NullToZero(rs.Fields("nevalidati").value)
             SumarUsers(5) = NullToZero(rs.Fields("users").value)
             SumarUsers(8) = NullToZero(rs.Fields("lockedusers").value)  
  end select
  rs.movenext 
 loop

 rs.Close
 set rs=nothing
 set cmd=nothing
 
 if Err.number<>0 then
	for i=0 to UBound(SumarUsers) 
	  SumarUsers(i)=-1
	next
	Err.Clear
 end if

 GetSumarUsers = SumarUsers    
End Function


' Functia intoarce un RecordSet ce contine lista cu utilizatori validati in functie de tipul acestora
' tipUser = tip-ul utilizatorilor (A,P,S).
Function GetValidatedUsers(tipUser, objCon)
 Dim StoredProcName
 
 select case UCase(tipUser)
  case "A" StoredProcName = "GetAdminsInfo"
  case "P" StoredProcName = "GetProfsInfo"
  case "S" StoredProcName = "GetStudsInfo"
 end select
 
 set cmd=Server.CreateObject("ADODB.Command")
 Set cmd.ActiveConnection = objCon
 cmd.CommandText = StoredProcName
 cmd.CommandType = adCmdStoredProc
 set rs = cmd.Execute
 set GetValidatedUsers = rs
 set cmd=nothing
End Function


' Blocheaza/Deblocheaza o lista de utilizatori.
' Intoarce true pentru succes altfel false.
Function LockUsers(usersList, locked, objCon)
 Const SQLLock = "UPDATE TBUsers SET locked=@1 WHERE id_user IN (@2)"
 
 On Error Resume Next
 objCon.Execute Replace(Replace(SQLLock, "@2", usersList), "@1", locked)
 If Err.number<>0 then
  Err.Clear 
  LockUsers = false
 Else
  LockUsers = true 
 End If
End Function


' Sterge inregistrarile utilizatorilor specificati prin usersList
' Intoarce true pentru succes
Function DeleteUsers(usersList, objCon)
 Const SQLDelete = "DELETE FROM TBUsers WHERE id_user IN (@1)"
 
 On Error Resume Next
 objCon.Execute Replace(SQLDelete, "@1", usersList)
 If Err.number<>0 then
  Err.Clear 
  DeleteUsers = false
 Else
  DeleteUsers = true
 End If  
End Function


' Modifica datele utilizatorului cu ID-ul dat prin userID
' PersData este un array ce contine noile valori
'   PersData(0) = nume
'   PersData(1) = prenume
'   PersData(2) = email
'   PersData(3) = telefon
'   PersData(4) = language
' Functia intoarce true pentru succes
Function UpdateUser(userID, ByVal PersData, objCon)
	Const SQL_UPDATE1 = "UPDATE TBPersons SET nume='@1', prenume='@2', email='@3', telefon='@4' WHERE id_person=@5"
	Const SQL_UPDATE2 = "UPDATE TBUsers SET id_lang=@lang WHERE id_user=@user"
	Dim idpers
	
	On Error Resume Next

	With GetUserByID(CLng(userID),objCon)
		idpers = .Fields("TBPersons.id_person").Value
	End With

	objCon.BeginTrans
	objCon.Execute Replace(Replace(Replace(Replace(Replace(SQL_UPDATE1,"@5",idpers),"@4",RemoveTags(PersData(3), 20)),"@3",RemoveTags(PersData(2),50)),"@2",RemoveTags(PersData(1),20)),"@1",RemoveTags(PersData(0),20))
	objCon.Execute Replace(Replace(SQL_UPDATE2, "@user", CStr(userID)), "@lang", CStr(PersData(4)))
	
	If Err.number<>0 then
		objCon.RollbackTrans 
		Err.Clear
		UpdateUser = false
	Else
		objCon.CommitTrans
		UpdateUser = true   
	End If
End Function


' Schimba neconditionat parola unui utilizator specificat prin userID
' Intoarce: 0 - schimbare reusita
'           1 - Err: parola nu se poate schimba cu sirul vid
'           2 - Err: nu se poate face modificarea in BD
Function ChangeUserPass1(userID, NewPass, objCon)
 Const SQL_CHPASS1 = "UPDATE TBUsers SET pass='@1' WHERE id_user=@2"

 On Error Resume Next

 If NewPass = "" then
   ChangeUserPass1 = 1
   Exit Function
 End If
 
 objCon.Execute Replace(Replace(SQL_CHPASS1,"@2",userID),"@1",NewPass)
 
 If Err.number<>0 then
   Err.Clear 
   ChangeUserPass1 = 2
 Else
   ChangeUserPass1 = 0
 End If
End Function


' Schimba conditionat parola unui utilizator. Vechea parola aflata in BD 
' trebuie sa coincida cu parola specificata prin parametrul OldPass
' Intoarce: 0 - schimbare reusita
'           1 - Err: parola nu se poate schimba cu sirul vid
'           2 - Err: eroare la lucrul cu BD
'           3 - Err: noua parola nu coincide cu vechea parola
Function ChangeUserPass2(userID, NewPass, OldPass, objCon)
  Dim OldPassDB

  On Error Resume Next
  Set rsuser = GetUserByID(CLng(userID),objCon)
  OldPassDB = rsuser.Fields("pass").Value
  rsuser.Close
  set rsuser=nothing

  If Err.number<>0 then
    ChangeUserPass2 = 2
    Exit Function
  End If

  If OldPass<>OldPassDB then
    ChangeUserPass2 = 3
    Exit Function
  End If

  ChangeUserPass2 = ChangeUserPass1(userID, NewPass, objCon)
End Function
%>
