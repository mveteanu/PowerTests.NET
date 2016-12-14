<%
' ******************************************************
' Array helpers
' ******************************************************

' Returns true if v is an element of ar
Function InArray(v,ar)
 Dim re
 re = false

 If Join(ar)<>"" then
	for i=0 to UBound(ar)
	 if ar(i) = v then
	   re = true
	   exit for
	 end if
	next
 End If
 
 InArray = re
End Function


' Add an element to an array
Function AddItemToArray(ByVal ar, arit)
 Redim Preserve ar(UBound(ar)+1)
 If IsObject(arit) then set ar(UBound(ar)) = arit else ar(UBound(ar)) = arit
 AddItemToArray = ar
End Function


' Deletes an element from an array
Function DelItemFromArray(ByVal ar, arit)
 Dim re()
 Dim i
 
 Redim re(-1)
 For Each i in ar
  If i<>arit then 
    Redim Preserve re(UBound(re)+1)
    re(UBound(re)) = i
  End If
 Next
 DelItemFromArray = re
End Function


' Returns an array with all elements that are in biga but not in littlea
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


' Returns ar1 + ar2 (without duplicates)
Function AddArrays(ar1, ar2)
 Dim re, it
 Redim re(-1)
 
 For Each it in ar1
  If not InArray(it,re) then re = AddItemToArray(re, it)
 Next
 For Each it in ar2
  If not InArray(it,re) then re = AddItemToArray(re, it)
 Next

 AddArrays = re 
End Function


' Returns the common elements of the 2 specified arrays
Function IntersectArrays(ar1, ar2)
 Dim re
 
 Redim re(-1)
 For each i in ar1
  If InArray(i,ar2) then re = AddItemToArray(re,i)
 Next
 
 IntersectArrays = re
End Function


' Returns an array of nrit items by randomly selecting
' elements from arr
Function GetRandomItems(arr, nrit)
 Dim re
 Dim maxindex, contor, rndnr,loopcount
 
 Redim re(-1)
 On Error Resume Next
 maxindex = UBound(arr)+1
 If Err.number<>0 then
  maxindex = 0
  Err.Clear 
 End If
 contor   = 0
 Randomize
 
 If nrit > maxindex then loopcount = maxindex else loopcount = nrit
 do until contor = loopcount
   rndnr = Int(maxindex * Rnd)
   If not InArray(arr(rndnr),re) then 
     re = AddItemToArray(re, arr(rndnr))
     contor = contor + 1
   End If  
 loop 

 GetRandomItems = re
End Function
%>