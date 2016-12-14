<%
' Embed a TDC control in the generated page
Public Sub AddTDC(strID, strDelim, strURL)
	If strDelim = Chr(1) Then strDelim = "&#01;"
	With Response
		.Write "<OBJECT id='"& strID &"' CLASSID='clsid:333C7BC4-460F-11D0-BC04-0080C7055A83' VIEWASTEXT>" & vbCrLf
		.Write "<PARAM NAME='UseHeader' VALUE='True'>" & vbCrLf
		.Write "<PARAM NAME='FieldDelim' VALUE='"& strDelim &"'>" & vbCrLf
		.Write "<PARAM NAME='DataURL' VALUE='"& strURL &"'>" & vbCrLf
		.Write "</OBJECT>" & vbCrLf & vbCrLf
	End With
End Sub


' Returns the HTML code for a SELECT control
' Dictionary Dict is used for item creation
' SelectID specifies the selected item
Function GetFillSelectFromDict(Dict, SelectID)
 Const OptionMach = "<OPTION @3 value='@1'>@2</OPTION>"
 Dim re, re1, it

 For each it in Dict.keys
  If CStr(SelectID) = CStr(it) then
    re1 = Replace(OptionMach, "@3", "SELECTED")
  Else
    re1 = Replace(OptionMach, "@3", "")  
  End If  
  re  = re & Replace(Replace(re1, "@2", Dict.item(it)), "@1", it) & vbCrLf
 Next
 
 GetFillSelectFromDict = re
End Function


' Adds a simple date selection control using 3 ComboBox-es
Sub CreateSimpleDateSelector(minyear, maxyear, left, top, ctrlname)
 Const LabelZi = "<span unselectable='on' class=TLabel style='left:@leftpx;top:@toppx;width:36px;height:13px;'>Day</span>"
 Const LabelLn = "<span unselectable='on' class=TLabel style='left:@leftpx;top:@toppx;width:36px;height:13px;'>Month</span>"
 Const LabelAn = "<span unselectable='on' class=TLabel style='left:@leftpx;top:@toppx;width:36px;height:13px;'>Year</span>"
 Const ComboBx = "<select unselectable='on' id='@id' class=TComboBox style='width:@widthpx;left:@leftpx;top:@toppx;'><option value=''></option>"
 Dim NumeLuni
 
 With Response
  .Write Replace(Replace(LabelZi, "@top", CStr(top+8)), "@left", CStr(left+176)) & vbCrLf
  .Write Replace(Replace(LabelLn, "@top", CStr(top+8)), "@left", CStr(left+72)) & vbCrLf
  .Write Replace(Replace(LabelAn, "@top", CStr(top+8)), "@left", CStr(left+8)) & vbCrLf
  
  .Write Replace(Replace(Replace(Replace(ComboBx, "@top", CStr(top+24)), "@left", CStr(left+176)), "@width", "57"), "@id", ctrlname & "zi")
  For i = 1 to 31
   .Write "<option value='" & CStr(i) & "'>" & CStr(i) & "</option>"
  Next 
  .Write "</select>" & vbCrLf

  .Write Replace(Replace(Replace(Replace(ComboBx, "@top", CStr(top+24)), "@left", CStr(left+72)), "@width", "97"), "@id", ctrlname & "ln")
  NumeLuni = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
  For i=1 to 12
   .Write "<option value='" & CStr(i) & "'>" & NumeLuni(i-1) & "</option>"  
  Next
  .Write "</select>" & vbCrLf

  .Write Replace(Replace(Replace(Replace(ComboBx, "@top", CStr(top+24)), "@left", CStr(left+8)), "@width", "57"), "@id", ctrlname & "an")
  For i = minyear to maxyear
   .Write "<option value='" & CStr(i) & "'>" & CStr(i) & "</option>"
  Next 
  .Write "</select>" & vbCrLf
 End With
End Sub


' Intoarce un string cu OPTION-uri pentru umplut un control de tipul SELECT
' La intrare: 
'   strTemplate  = string template cu formatul option-ului ce se doreste intors... 
'                  Ex: <option @SELECTED id=@1 value=@2>@3 / @4</option>
'   strSelTemplate = string template ce va fi evaluat pentru a se decide daca se pune sau nu SELECTED
'				   Ex: "@0>=10"
'   arStoredProc = 1) string cu numele procedurii stocate/tabelului, sau ...
'                = 2) array in formatul: index0 = nume procedura stocata, index1...n = parametrii query-ului
'   arFields     = array cu field-urile din recordset-ul obtinut folosite pentru completarea template-ului
'                  Acestea pot fi specificate numeric prin pozitie sau string prin numele campului
'   objCon       = 1) obiect Connection existent, sau ...
'                = 2) Nothing daca se doreste crearea in interior a obiectului Connection
Function GetOptionsForSelect(ByVal strTemplate, ByVal strSelTemplate, arStoredProc, arFields, objCon)
	Dim strStoredProc, arStoredProcParams, bCreateCon
	Dim objInternalCon, objInternalRS, objInternalCmd
	Dim reAll, reLine
	
	reAll = ""
	bCreateCon = CBool(objCon Is Nothing)
	If IsArray(arStoredProc) Then
		strStoredProc = arStoredProc(0)
		If UBound(arStoredProc) > 0 Then
			Redim arStoredProcParams(UBound(arStoredProc)-1)
			For i = 0 To UBound(arStoredProcParams)
				arStoredProcParams(i) = arStoredProc(i+1) 
			Next
		End If
	Else
		strStoredProc = arStoredProc
	End If
	
	If bCreateCon Then 
		Set objInternalCon = Server.CreateObject("ADODB.Connection")
		objInternalCon.Open Application("DSN")
	Else
		Set objInternalCon = objCon
	End If
	
	If IsEmpty(arStoredProcParams) Then
		Set objInternalRS = objInternalCon.Execute(strStoredProc)
	Else
		Set objInternalCmd = Server.CreateObject("ADODB.Command")
		Set objInternalCmd.ActiveConnection = objInternalCon
		objInternalCmd.CommandText = strStoredProc
		objInternalCmd.CommandType = adCmdStoredProc
		Set objInternalRS = objInternalCmd.Execute(,arStoredProcParams)
		Set objInternalCmd = Nothing
	End If
	
	If Len(strSelTemplate) = 0 Then 
		strTemplate = Replace(strTemplate, "@SELECTED", "")
	Else
		For i=0 To UBound(arFields)
			If IsNumeric(arFields(i)) Then
				s = "objInternalRS.Fields(" & arFields(i) & ").Value"
			Else
				s = "objInternalRS.Fields(""" & arFields(i) & """).Value"
			End If
			strSelTemplate = Replace(strSelTemplate, "@" & i, s)
		Next
	End If
	Do Until objInternalRS.EOF
		reLine = strTemplate
		If strSelTemplate<>"" Then
			If Eval(strSelTemplate) Then
				reLine = Replace(strTemplate, "@SELECTED", "SELECTED")
			Else
				reLine = Replace(strTemplate, "@SELECTED", "")
			End If
		End If
		For i=0 To UBound(arFields)
			reLine = Replace(reLine, "@" & i, Nz(objInternalRS.Fields(arFields(i)).Value, ""))
		Next
		reAll = reAll & reLine & vbCrLf
		objInternalRS.MoveNext
	Loop
	
	objInternalRS.Close
	Set objInternalRS = Nothing
	If bCreateCon Then 
		objInternalCon.Close
		Set objInternalCon = Nothing
	End If
	
	GetOptionsForSelect = reAll
End Function
%>
