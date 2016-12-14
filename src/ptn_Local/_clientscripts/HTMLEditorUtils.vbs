'<script language="vbscript">

Dim FileNumber
Dim DSel
FileNumber = 0

' Returns true if v is found in array ar
Function InArray(v,ar)
	Dim re

	re = False
 
	If Join(ar)<>"" Then
		For i=0 to UBound(ar)
			If ar(i) = v Then
				re = True
				Exit For
			End If
		Next
	End If
 
	InArray = re
End Function


' Returns an array with elements that are in biga but not in littlea
Function ArrayDif(biga, littlea)
	Dim re()
	Dim nri
 
	If Join(biga)<>"" then
		nri = 0
		for i=0 to UBound(biga)
			If not InArray(biga(i),littlea) then
				Redim Preserve re(nri)
				re(nri) = biga(i)
				nri = nri + 1
			end if
		next
	End If
 
	ArrayDif = re
End Function


' Decode an Url by replacing the %xx sequences with corresponding chars
' The + sign is not translated to " " due to the way IE handles this char
' Sequences of type %uxxxx are not translated !
Function URLDecode(Expression)
	Dim strSource, strTemp, strResult
	Dim lngPos, s

	strSource = Expression

	For lngPos = 1 To Len(strSource)
		strTemp = Mid(strSource, lngPos, 1)
		If strTemp = "%" Then
			If lngPos + 2 < Len(strSource) Then
				s = Mid(strSource, lngPos + 1, 2)
				If IsNumeric("&H" & s) then 
					strResult = strResult & Chr(CInt("&H" & s))
				Else
					strResult = strResult & "%" & s
				End If   
				lngPos = lngPos + 2
			End If
		Else
			strResult = strResult & strTemp
		End If
	Next
	URLDecode = strResult
End Function


' Reads the value of an query string parameter
' The function is equivalent with ASP method: Server.RequestQueryString
Function QueryString(url, item)
 Dim p1, p2
 Dim re
 
 p1 = InstrRev(url,item & "=") + Len(item & "=")
 if p1 = Len(item & "=") then 
   re = ""
 else  
   p2 = Instr(p1,url,"&")
   if p2 = 0 then p2 = Len(url) + 1
   re = Mid(url,p1,p2-p1)
 end if 
  
 QueryString = re
End Function


' Returns an array with SRC attribute of all images located
' in the EditPage container
Function GetEditPageImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then 
     Redim Preserve im(nri)
     im(nri) = i.src
     nri = nri + 1
   end if
 next
 
 GetEditPageImages = im
End Function


' Returns an array with SRC attribute of all images located
' in the EditPage container and which are comming from the server
' using http:// protocol (the server is not checked, though;-))
Function GetEditPageServerImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then
     if LCase(Left(i.src,7)) = "http://" then 
       Redim Preserve im(nri)
       im(nri) = i.src
       nri = nri + 1
     end if
   end if  
 next
 
 GetEditPageServerImages = im
End Function



' Returns an array with SRC attribute of all images located
' in the EditPage container and which are comming from a local or UNC path
Function GetEditPageLocalImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then
     if LCase(Left(i.src,7)) = "file://" then 
       Redim Preserve im(nri)
       im(nri) = i.src
       nri = nri + 1
     end if
   end if  
 next
 
 GetEditPageLocalImages = im
End Function


' Convert from a local file Url format to a local file path format
' Example: From file:///d:/t/img/... or file://vma-athlon/t/img/...
' to: d:\t\img\... or \\vma-athlon\t\img\...
' If the input string doesn't follow local file Url format then
' it is return unchanged
Function LocalURLToFileName(fileurl)
 Dim s
 Dim re
 
 s = Replace(fileurl,"/","\")
 if LCase(Left(s,8))="file:\\\" then 
   re = URLDecode(Right(s,Len(s)-8))
 elseif LCase(Left(s,7))="file:\\" then
   re = URLDecode("\\" & Right(s,Len(s)-7))
 elseif (UCase(Left(s,1))>="A") and (UCase(Left(s,1))<="Z") and (Mid(s,2,2)=":\") then
   re = URLDecode(s)
 else
   re = fileurl
 end if
 
 LocalURLToFileName = re
End Function


' Extracts the values specified by idname from an array of Urls
' and concats them in a CSV string
Function FileNamesArrayToIDCSV(ar, idname)
 Dim re
 Dim qs
 
 re = ""
 If Join(ar)<>"" then
  For i = 0 to UBound(ar)
   qs = QueryString(ar(i),idname)
   If qs<>"" then re = re & qs & ","
  Next
  re = Left(re,Len(re)-1)
 End If
 
 FileNamesArrayToIDCSV = re
End Function


' This sub is executed by pressing the Insert Picture button.
' The effect is displaying of the miniform for typing the file name
Sub ShowAddFileForm
 set DSel = document.selection.createRange 
 FileNumber = FileNumber + 1
 ChosePictureFormular.insertAdjacentHTML "beforeEnd",  "<INPUT type='file' name='File" & CStr(FileNumber) & "' id='File" & CStr(FileNumber) & "'>"
 set ifile = ChosePictureForm.all("File" & CStr(FileNumber))
 with ifile
  .className = "TEdit"
  .style.left = 20
  .style.top = 40
  .style.width = 250
  .style.height = 21
  ChosePictureForm.style.visibility = ""
  .focus
 end with
 set ifile = nothing
End Sub


' Is executed automatically when the miniform's OK button is pressed
Sub ChosePictureFormButOK_onclick
 Dim im 
 im = ChosePictureForm.all("File" & CStr(FileNumber)).Value
 if im<>"" then
   if LCase(document.selection.type) <> "control" then 
     DSel.select
     document.execCommand "InsertImage", false, im
   end if  
 end if
End Sub


' Removes from the hidden files form those INPUT FILE elements
' that point to images removed from the document during editing.
' Also removes duplicated entries.
Sub CleanFilesForm(fform, docimg)
 Dim este, fty
 
 for each ff in fform.elements
  if ff.type = "file" then
	este = false
	If Join(docimg)<>"" then
		for li = 0 to UBound(docimg)
		  if ff.value = LocalURLToFileName(docimg(li)) then este = true
		next 
	End If 	
	if not este then fform.removeChild ff
  end if
 next
 ' Now removes duplicate entris for lowering network traffic
 ' and DB space requirements
 for i1 = 0 to fform.elements.length-1
   for i2 = i1+1 to fform.elements.length-1
     if (LCase(fform.elements(i1).type)="file") and (LCase(fform.elements(i2).type)="file") and (fform.elements(i1).value = fform.elements(i2).value) then
       fform.removeChild fform.elements(i1)
       Exit For
     end if
   next
 next
End Sub


' Returns as a CSV text all the images that should be uploaded to server
' The info is extracted from the hidden form
Function GetUploadFields(fform)
 Dim re
 
 re = ""
 For each ff in fform.elements
  if ff.type = "file" then
    re = re & CStr(ff.name) & ","
  end if  
 Next
 If Len(re)>0 then re = Left(re,Len(re)-1)
 
 GetUploadFields = re
End Function
