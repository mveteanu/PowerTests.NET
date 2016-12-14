'<script language="vbscript">

' Returns as CSV all questions IDs from the specified <SELECT>
' If onlyselected = true then only the selected ones will be returned
Function GetItemsFromSelect(selectid, onlyselected)
	Dim re, opt
  
	re = ""
	For Each opt in selectid.Options
		If onlyselected then 
			If opt.selected then
				re = re & opt.value & ","
			End If
		Else
			re = re & opt.value & ","
		End If    
	Next
  
	If re<>"" then re = Left(re, Len(re)-Len(","))
	GetItemsFromSelect = re
End Function


' Returns the position of an element form a SELECT
Function GetSelectElement(selectid, elem)
	Dim re, i
 
	re = -1
	For i = 0 to selectid.options.length
		If selectid.options(i).value = elem then
			re = i
			Exit For
		End If
	Next
 
	GetSelectElement = re
End Function


' Removes all OPTION elements from a SELECT
Function ClearSelect(objSel)
	Dim i
	
	For i = 0 To objSel.Options.Length - 1
		objSel.Options.Remove(0)
	Next
End Function


' Remove specified item from a select
' Returns the value of removed item or empty string if none is removed
Function RemoveItemFromSelect(ByRef objSel, optIndex)
	Dim re

	re = ""
	If optIndex > -1 Then
		re = objSel.options(optIndex).Value
		objSel.options.remove(optIndex)
		If optIndex > objSel.options.length - 1 Then
			objSel.selectedIndex = 0
		Else
			objSel.selectedIndex = optIndex
		End If
	End If
	RemoveItemFromSelect = re
End Function


' Removes the item with specified value from a SELECT
Sub RemoveItemFromSelectByValue(ByRef objSel, optValue)
	For i = 0 To objSel.Options.Length-1
		If objSel.options(i).value = optValue Then
			Call RemoveItemFromSelect(objSel, i)
			Exit For
		End If
	Next
End Sub


' Add a new OPTION item to a SELECT
Sub AddItemToSelect(objSel, strValue, strText)
	Dim newel

	Set newel = document.createElement("OPTION")
	objSel.Options.Add newel
	newel.innerText = strText
	newel.Value     = strValue
	Set newel = nothing
End Sub



' Fills a SELECT from a filtered recordset
' If strColFiltVal = intAllVal then all elements will be included
Function FillSelectFromFilteredRS(objSel, objRS, strColVal, strColText, strColFilt, strColFiltVal, strAllVal)
	Dim re, canadd

	re = 0
	If not (objRS.BOF and objRS.EOF) Then 
		objRS.MoveFirst
		Do While Not objRS.EOF
			canadd = false
			If CStr(strColFiltVal) = CStr(strAllVal) then 
				canadd = true
			Else
				If CStr(objRS.Fields(strColFilt).Value) = CStr(strColFiltVal) Then canadd = true
			End If
			If canadd then
				Call AddItemToSelect(objSel, objRS.Fields(strColVal).Value, objRS.Fields(strColText).Value)
				re = re + 1
			End If
			objRS.MoveNext
		Loop
	End If
	FillSelectFromFilteredRS = re
End Function


' Fills a combobox using the data from a TDC.
' Supports adding of (all) and (none) items
Sub QuickFillCombo(objSel, objTDC, FldVal, FldTxt, TxtAll, TxtNone)
	Call ClearSelect(objSel)
	If Len(TxtAll)>0 Then Call AddItemToSelect(objSel, -1, TxtAll)
	If Len(TxtNone)>0 Then Call AddItemToSelect(objSel, 0, TxtNone)
	Call FillSelectFromFilteredRS(objSel, objTDC.Recordset, FldVal, FldTxt, "", -1, -1)
End Sub


' This is used by the 4 buttons (>, >>, <, <<) 
' from the multichoice control
Sub MoveItemsBetweenSelects(FromSel, ToSel, IDList)
	Dim it, idar, oldpos
	Dim newel
 
	idar = Split(IDList, ",", -1, 1)
	For Each it in idar
		oldpos = GetSelectElement(FromSel, it)
		If oldpos<>-1 then
			Set newel = document.createElement("OPTION")
			ToSel.Options.Add newel
			newel.innerText = FromSel.Options(oldpos).Text
			newel.Value     = FromSel.Options(oldpos).Value
			FromSel.Options.Remove(oldpos)
			Set newel = nothing
		End If
	Next
End Sub


' This is used by the 4 buttons (>, >>, <, <<) 
' from the multichoice control
Sub HandleButtonsBetweenSelects(objListBox1, objListBox2, strMoveCmd)
	Select Case strMoveCmd
		case ">"  MoveItemsBetweenSelects objListBox1, objListBox2, GetItemsFromSelect(objListBox1, true)
		case ">>" MoveItemsBetweenSelects objListBox1, objListBox2, GetItemsFromSelect(objListBox1, false)
		case "<"  MoveItemsBetweenSelects objListBox2, objListBox1, GetItemsFromSelect(objListBox2, true)
		case "<<" MoveItemsBetweenSelects objListBox2, objListBox1, GetItemsFromSelect(objListBox2, false)
	End Select
End Sub

