<%

' Sends an Email
Sub SendEmail(msgTo, msgFrom, msgSubject, msgBody) 
'
'Response.Write "To: "&msgTo&"<br>"
'Response.Write "From: "&msgFrom&"<br>"
'Response.Write "Subject: "&msgSubject&"<br>"
'Response.Write "Body: "&msgBody&"<br>"
'
End Sub



' Sends the account activation email to a user 
' Email will come from system administrator
' tipemail 
'          = 1 Account activated email
'          = 2 Opening account denied email
' Returns true for success
Function SendConfirmationEmail(iduser, idadmin, tipemail, objCon)
  Const Msg_AcceptaCerere  = "Your enrolment request was accepted. Please use your login information to connect to PTN system."
  Const Msg_RespingeCerere = "Your enrolment request was denied by a PTN administrator. Please contact us for more information."
  Dim BodyMessage
  
  select case tipemail
    case 1 BodyMessage = Msg_AcceptaCerere
    case 2 BodyMessage = Msg_RespingeCerere
  end select
  
  on error resume next
  set rsuser  = GetUserByID(iduser,objCon)
  set rsadmin = GetUserByID(idadmin,objCon)

  SendEmail rsuser.Fields("email").value, _
            rsadmin.Fields("email").value, _
            "PowerTests .NET notification", _
            BodyMessage

  rsuser.Close
  rsadmin.Close
  set rsuser  = nothing
  set rsadmin = nothing
  if Err.number<>0 then
    Err.Clear 
    SendConfirmationEmail = false
  else
    SendConfirmationEmail = true
  end if
End Function
%>