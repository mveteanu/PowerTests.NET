'<script language="vbscript">

' Copy the values contained in fields of a form to fields in other form
' Returns True in case of success
' Note: Form elements must have id's and have a correspondent in the other form
Function CopyForms(f1,f2)
	On Error Resume Next
	For i = 0 to f1.all.length-1
		If f1.all(i).tagname = "INPUT" or _
			f1.all(i).tagname = "TEXTAREA" or _
			f1.all(i).tagname = "SELECT" Then 
				If f1.all(i).type<>"button" Then f2.all(f1.all(i).id).value = f1.all(i).value
		End If   
	next
	If Err.number<>0 Then
		CopyForms = False
		Err.Clear
	Else
		CopyForms = True
	End If   
End Function


' Compares two forms, element by element
' and returns true if they are identical
Function CompareForms(f1,f2)
	On Error Resume Next
	Dim resp
	resp = True
	For i = 0 To f1.all.length-1
		If f1.all(i).tagname = "INPUT" or _
			f1.all(i).tagname = "TEXTAREA" or _
			f1.all(i).tagname = "SELECT" Then
				If f1.all(i).type<>"button" Then If f2.all(f1.all(i).id).value <> f1.all(i).value Then resp = False
		End If   
	Next

	If Err.number<>0 Then
		CompareForms = False
		Err.Clear
	Else  
		CompareForms = resp
	End If  
End Function
