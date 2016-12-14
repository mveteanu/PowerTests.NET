'<script language="vbscript">

 ' Returns as an array the value of ITEMs from a combobox
 Function GetComboItems(comboid)
  Dim re()
  Dim nri, i
  
  nri = comboid.Options.Length
  If nri = 0 then Exit Function
  
  ReDim re(nri)
  For i=0 to nri-1
    re(i) = comboid.Options(i).Value
  Next
  
  GetComboItems = re
 End Function


 ' Fill a ComboBox with values from an array
 Sub SetComboItems(comboid, itemsar)
  Dim i, it
  
  If not IsArray(itemsar) Then Exit Sub
  
  comboid.innerHTML = ""
    
  For Each i in itemsar
   set it = document.createElement("OPTION")
   it.text  = CStr(i)
   it.value = CStr(i)
   comboid.Add it
   set it = nothing
  Next
 End Sub

 
 ' This is invoked by pressing the Edit link located near the ComboBox controls
 Sub EditCombo(comboid)
  window.event.returnValue = false
  SetComboItems comboid, ShowModalDialog("pbeditcombo.asp", GetComboItems(comboid) , "dialogWidth=200px;dialogHeight=200px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
 End Sub


 ' Returns a number containing (in packed format) the properties of an Edit
 ' Format: b5 b4 b3 b2 b1 b0
 ' b0 = 0 numeric answer
 '    b3 b2 b1 = precision
 '    b4 = ingnore non-numeric characters
 ' b0 = 1 text answer
 '    b3 b2 b1 = matching method
 '    b4 = case sensitive
 '    b5 = ignore trailing spaces
 Function GetEditProp(propeditid)
   Dim re
   
   If propeditid.Value="" then
     re = 20
   Else
     re = CInt(propeditid.Value)
   End If
   
   GetEditProp = re  
 End Function

 
 ' Sets the properties of an Edit
 Sub SetEditProp(propeditid, newprop)
   If CStr(newprop)<>"" then
     propeditid.Value = newprop
   End If  
 End Sub


 ' This is invoked by pressing the Properties link located near
 ' Edit type controls
 Sub EditProperties(editid)
  window.event.returnValue = false
  SetEditProp editid, ShowModalDialog("pbeditprop.asp", GetEditProp(editid) , "dialogWidth=384px;dialogHeight=250px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
 End Sub


 ' Adds an answer in the answering area
 Sub AddAnswer(ByVal AnswZone, ByRef NrAns, ByRef TipRasp)
  Const MachetaRadio = "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input id='answer@2' name=answer type=radio class='TAnswerButton' title='Bifati raspunsul corect'></span>"
  Const MachetaCheck = "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input id='answer@2' name=answer type=checkbox class='TAnswerButton' title='Bifati raspunsurile corecte'></span>"
  Const MachetaCombo = "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><select id='answer@2' name=answer rows=1 class='TAnswerComboBox' title='Selectati raspunsul corect'></select>&nbsp;&nbsp;<a href=# onclick='vbscript:EditCombo(answer@2)' unselectable='on' title='Editati lista cu variante de raspuns'>Edit</a></span>"
  Const MachetaEdit  = "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input id='answer@2' name=answer type=edit class='TAnswerEdit' title='Introduceti raspunsul corect'><input id='propanswer@2' type=edit value=20 style='display:none;'>&nbsp;&nbsp;<a href=# onclick='vbscript:EditProperties(propanswer@2)' unselectable='on' title='Setare proprietati raspuns'>Proprietati</a></span>"
  
  Dim Macheta
  
  If NrAns < 1 then 
    TipRasp = ShowModalDialog("pbrequestanstype.asp",, "dialogWidth=386px;dialogHeight=256px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no") 
  ElseIf NrAns > 25 then 
    Exit Sub
  end if

  select case TipRasp
    case 1 Macheta = MachetaRadio
    case 2 Macheta = MachetaCheck
    case 3 Macheta = MachetaCombo
    case 4 Macheta = MachetaEdit
    case else Exit Sub
  end select
  
  NrAns = NrAns + 1
  AnswZone.insertAdjacentHTML "beforeEnd", Replace(Replace(Macheta, "@1", Chr(NrAns+64)), "@2", CStr(NrAns))
 End Sub

 
 ' Deletes last added answer
 Sub RemoveAnswer(ByVal AnswZone, ByRef NrAns)
  If NrAns < 1 then Exit Sub
  Set ansr = AnsDIV.All("allanswer" & CStr(NrAns))
  AnswZone.RemoveChild ansr
  NrAns = NrAns - 1
 End Sub



 ' Returns in a packed format (multiline strings) the information
 ' from the answering area
 ' If controls were not filled correctly then the empty string is returned
 Function GetAnswers(ByVal AnswZone, ByVal NrAns, ByVal TipRasp)
   Dim i, s, s2, ob, b
   Dim v1, v2
   
   If NrAns < 1 then
    GetAnswers = ""
    Exit Function
   End If 
   
   s = CStr(TipRasp) & vbCrLf
   select case TipRasp
     case 1    b = false
               If NrAns > 1 then
				For i = 1 to NrAns
				  If AnswZone.All("answer" & CStr(i)).Checked then
				    s = s & CStr(CInt(true)) & vbCrLf
				    b = true
				  Else
				    s = s & CStr(CInt(false)) & vbCrLf
				  End If
				Next
               End If
               If b = false then s = ""
     case 2    For i = 1 to NrAns
                 If AnswZone.All("answer" & CStr(i)).Checked then
                   s = s & CStr(CInt(true)) & vbCrLf
                 Else
                   s = s & CStr(CInt(false)) & vbCrLf
                 End If  
               Next
     case 3    For i = 1 to NrAns
                 set ob = AnswZone.All("answer" & CStr(i))
                 If ob.Options.Length < 1 then
                   s = ""
                   Exit For
                 End If
                 s  = s & CStr(ob.selectedIndex) & Chr(2)
                 s2 = Join(GetComboItems(ob), Chr(3))
                 s  = s & Left(s2,Len(s2)-Len(Chr(3))) & vbCrLf
                 set ob = nothing
               Next
     case 4    For i = 1 to NrAns
                 v1 = AnswZone.All("answer" & CStr(i)).Value
                 v2 = AnswZone.All("propanswer" & CStr(i)).Value
                 If (CStr(v1)="") or (CStr(v2)="") then
                   s = ""
                   Exit For
                 End If
                 s = s & v1 & Chr(2) & v2 & vbCrLf
               Next
   end select
   
   If Len(s)>Len(vbCrLf) then
     GetAnswers = Left(s,Len(s)-Len(vbCrLf))
   Else
     GetAnswers = s 
   End If  
 End Function
