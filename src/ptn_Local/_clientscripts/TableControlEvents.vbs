'<script language="vbscript">

' BEGIN PUBLIC SECTION

' Returns a CSV will the IDs from selected rows
Public Function TableGetSelected(tableID)
	Dim mRow, sRet
	
	sRet = ""
	For Each mRow In tableID.rows 
		If mRow.cells(0).className = "TTableRowSelected" Then _
			sRet = sRet & mRow.Cells(0).children(0).innerText & ","
	Next
	If Len(sRet) > 0 Then sRet = Left(sRet, Len(sRet)-1)
	TableGetSelected = sRet
End Function


' Select/Unselect the rows depending on the state parameter
Public Sub TableSelectAll(tableID,state)
	Dim mRow, mCol
	For Each mRow In tableID.rows 
      For Each mCol In mRow.Cells
        if state = true then
          mCol.className = "TTableRowSelected" 
        else
          mCol.className = "TTableRowUnSelected"
        end if
      next  
	Next
End Sub


' BEGIN PRIVATE SECTION


Private Sub HandleTableClick(tipsel)
	Dim obj, sTag, mRow, mTabl

    Set mRow  = nothing
    Set mTabl = nothing
   
	Set obj=window.event.srcElement
	sTag = obj.tagName
	Set mRow = Nothing
	If sTag = "TD" Then 
	  set mRow  = obj.parentElement
	  set mTabl = obj.parentElement.parentElement.parentElement
	ElseIf sTag = "SPAN" Then
	  set mRow  = obj.parentElement.parentElement
	  set mTabl = obj.parentElement.parentElement.parentElement.parentElement
	End If
	
    If Not mTabl Is Nothing Then TableSwitchRow mTabl, mRow , tipsel
End Sub


Private Sub TableSwitchRow(objTabl, objRow, tipsel)
 Dim ir,ic
 
 If tipsel=1 then
   If objRow.Cells(0).className = "TTableRowSelected" Then Exit Sub
   For Each ir In objTabl.rows
     If ir.Cells(0).className = "TTableRowSelected" Then
       For Each ic In ir.cells
         ic.className = "TTableRowUnSelected"
       Next
     End If
   Next
   For Each ic In objRow.cells
     ic.className = "TTableRowSelected"
   Next
 Else
 	For Each ic In objRow.cells
		If ic.className = "TTableRowSelected" Then
			ic.className = "TTableRowUnSelected"
		ElseIf ic.ClassName = "TTableRowUnSelected" Then
			ic.className = "TTableRowSelected"
		End If
	Next
 End If	
End Sub


Private Sub HandleTableHeaderSort(tdcname, nrdiv, lastdiv)
 Dim obj, sTag, maindiv, sortk, sortname, i

 Set obj=window.event.srcElement
 sTag = obj.tagName

 if sTag="DIV" then
   sortk = obj.children(0).innerhtml
   sortname = obj.children(1).innerhtml
   set maindiv = obj.parentElement
 elseif sTag = "SPAN" Then
   sortk = obj.parentElement.children(0).innerhtml
   sortname = obj.parentElement.children(1).innerhtml
   set maindiv = obj.parentElement.parentElement
 end if
 for i=0 to lastdiv
   if i<>nrdiv then
     maindiv.children(i).children(0).style.display="none"
   else
     maindiv.children(i).children(0).innerhtml=CStr(CInt(sortk) xor 3)
     maindiv.children(i).children(0).style.display=""
   end if
 next 

 if sortk="6" then 
    sortk="+"
 else
    sortk="-"
 end if

 tdcname.sort=sortk & sortname
 tdcname.reset
End Sub


' Gets information from a TDC row. 
' Returns an array of the same size as ArrayOfFields
Function GetTDCData(objTDC, RowId, ArrayOfFields)
 Dim rs
 Dim re()
 
 Redim re(UBound(ArrayOfFields))
 Set rs = objTDC.RecordSet.Clone
 Do Until rs.EOF
  If CLng(rs.Fields("id").Value) = CLng(RowId) then
    For i = 0 to UBound(ArrayOfFields)
     re(i) = rs.Fields(ArrayOfFields(i)).Value
    Next
    Exit Do
  End If
  rs.MoveNext 
 Loop
 
 GetTDCData = re
End Function
