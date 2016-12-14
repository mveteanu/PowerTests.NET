<!-- #include file="utils.asp" -->
<!-- #include file="arrayutils.asp" -->
<!-- #include file="problems.asp" -->
<%
' *********************************************************************
' Server side support for test definition handling
' Data:  April 18, 2001
' *********************************************************************

Class PTNTestCategFilter
  Public CategID
  Public CategName
  Public Comparation
  Public NrProblems
End Class

Class PTNTestDefinition
  Private objCon

  Public TestID
  Public Name
  Public Comments
  Public Time
  Public MaxSustineri
  Public TstPublic
  Public PBIncluse      ' Dictionary 
  Public PBExcluse      ' Dictionary
  Public MaxRandom
  Public CategFilt      ' Dictionary
  
  Private Sub Class_Initialize()
    Set PBIncluse = CreateObject("Scripting.Dictionary")
    Set PBExcluse = CreateObject("Scripting.Dictionary")
    Set CategFilt = CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate()
    Set PBIncluse = nothing
    Set PBExcluse = nothing
    Set CategFilt = nothing
    Set objCon    = nothing
  End Sub
  
  Public Sub LoadTestDefinition(tstid, objConect)
    Const SQLSelTst = "SELECT * FROM TBTests WHERE id_test=@1"
    Dim   myCmd, rs, catfil, contor
    
    TestID     = tstid
    Set objCon = objConect
    
    Set rs = objCon.Execute(Replace(SQLSelTst,"@1",CStr(TestID)))
    
    If not rs.EOF then
       Name          = rs.Fields("numetest").Value
       Comments      = rs.Fields("comentarii").Value
       Time          = rs.Fields("timp").Value
       MaxSustineri  = rs.Fields("nrmaxsustineri").Value
       TstPublic     = rs.Fields("testpublic").Value
       MaxRandom     = rs.Fields("maxrandom").Value
    End If
    If rs.State = adStateOpen then rs.Close
    Set rs = nothing
    
    Set myCmd = Server.CreateObject("ADODB.Command")
    Set myCmd.ActiveConnection = objCon
    myCmd.CommandText = "GetTstExplicitPbs"
    myCmd.CommandType = adCmdStoredProc
    Set rs = myCmd.Execute(,TestID)
    do while not rs.EOF
      Select case rs.Fields("filter").value
        case 0 PBIncluse.Add CStr(rs.Fields("id_problem").value), rs.Fields("numeproblema").value
        case 1 PBExcluse.Add CStr(rs.Fields("id_problem").value), rs.Fields("numeproblema").value
      End Select
      rs.MoveNext  
    loop
    If rs.State = adStateOpen then rs.Close
    Set rs    = nothing

    myCmd.CommandText = "GetTstFiltredCategs"
    Set rs = myCmd.Execute(,TestID)
    contor = 0
    do while not rs.EOF
      set catfil = New PTNTestCategFilter
      catfil.CategID     = rs.Fields("id_categpb").value
      catfil.CategName   = rs.Fields("numecateg").value
      catfil.Comparation = rs.Fields("filter").value
      catfil.NrProblems  = rs.Fields("nrproblems").value
      CategFilt.Add CStr(contor), catfil
      set catfil = nothing
      rs.MoveNext
      contor = contor + 1
    loop
    If rs.State = adStateOpen then rs.Close
    Set rs    = nothing
    Set myCmd = nothing
  End Sub

End Class


Class PTNTestCategDetails
 Public Name
 Public ProblemsArray
 Public ProblemsWanted
 Public ProblemsLoaded
End Class


Class PTNTestGenerator
 Public CursID
 Public GenerationDate   ' Test generation date
 Public TestDefinition   ' Test Definition of type PTNTestDefinition
 Public AllProblemsArray
 Public IncludePbArray
 Public ExcludePbArray
 Public FixCategsArray
 Public MinCategsArray
 Public MaxCategsArray
 
 Private Sub Class_Initialize()
  CursID = CLng(Session("CursID"))
  Redim AllProblemsArray(-1)
  Redim FixCategsArray(-1)
  Redim MinCategsArray(-1)
  Redim MaxCategsArray(-1)
  Set TestDefinition = New PTNTestDefinition
 End Sub

 Private Sub Class_Terminate()
  Set TestDefinition = nothing
 End Sub


 Public Sub LoadTest(tstid, objConect)
  Dim i,ar,arc, ar1, cdet
  Dim myCmd, rs
  
  ' Load test definition
  TestDefinition.LoadTestDefinition tstid, objConect

  ' Creates an array with all questions from a course
  AllProblemsArray = GetTotalPbArray(CursID, objConect)
  
  ' Creates the array with the questions that MUST be included
  Redim IncludePbArray(TestDefinition.PBIncluse.Count-1)
  ar  = TestDefinition.PBIncluse.Keys
  arc = TestDefinition.PBIncluse.Count
  For i = 0 To arc-1: IncludePbArray(i) = CLng(ar(i)): Next
  
  ' Creates the array with the questions that MUST NOT be included
  Redim ExcludePbArray(TestDefinition.PBExcluse.Count-1)
  ar  = TestDefinition.PBExcluse.Keys
  arc = TestDefinition.PBExcluse.Count
  For i = 0 To arc-1: ExcludePbArray(i) = CLng(ar(i)): Next
  
  ' Builds the arrays with questions from categories
  ar1 = TestDefinition.CategFilt.Items
  arc = TestDefinition.CategFilt.Count
  For i = 0 To arc-1
   Set cdet = New PTNTestCategDetails
   cdet.Name           = ar1(i).CategName
   cdet.ProblemsWanted = ar1(i).NrProblems
   cdet.ProblemsLoaded = 0
   cdet.ProblemsArray  = GetCategPbArray(ar1(i).CategID, objConect)
   Select Case ar1(i).Comparation
     case 0 FixCategsArray = AddItemToArray(FixCategsArray, cdet)
     case 1 MinCategsArray = AddItemToArray(MinCategsArray, cdet)
     case 2 MaxCategsArray = AddItemToArray(MaxCategsArray, cdet)
   End Select
   Set cdet = nothing
  Next
 End Sub


 Public Function GetGeneratedTest
  Dim AvailableProblems, GeneratedTestArray
  Dim TempPbs, NrAvailPbs
  
  AvailableProblems = AllProblemsArray
  Redim GeneratedTestArray(-1)
  For Each i In MaxCategsArray: i.ProblemsLoaded = 0: Next
  
  ' Adds the questions from IncludePb list and 
  ' remove from AvailableProblems the questions that are in IncludePb and ExcludePb
  GeneratedTestArray  = IncludePbArray
  AvailableProblems   = ArrayDif(AvailableProblems, IncludePbArray)
  AvailableProblems   = ArrayDif(AvailableProblems, ExcludePbArray)
 
  ' Adds a random number of questions from FixProblems categories and
  ' remove from AvailableProblems the entire category FixProblems
  For Each i In FixCategsArray
   GeneratedTestArray = AddArrays(GeneratedTestArray, GetRandomItems(ArrayDif(ArrayDif(i.ProblemsArray,GeneratedTestArray),ExcludePbArray), i.ProblemsWanted))
   AvailableProblems  = ArrayDif(AvailableProblems, i.ProblemsArray)

   For Each j1 In MaxCategsArray
    For Each j2 In i.ProblemsArray
     If InArray(j2,j1.ProblemsArray) then j1.ProblemsLoaded = j1.ProblemsLoaded + 1
    Next 
   Next

  Next

  ' Se adauga in mod aleator un nr. de probleme specificat de categoriile MinProblems
  ' Se elimina din AvailableProblems doar problemele adaugate din respectivele categorii
  For Each i In MinCategsArray
   TempPbs = GetRandomItems(ArrayDif(ArrayDif(i.ProblemsArray,GeneratedTestArray),ExcludePbArray), i.ProblemsWanted)
   GeneratedTestArray = AddArrays(GeneratedTestArray, TempPbs)
   AvailableProblems  = ArrayDif(AvailableProblems, TempPbs)

   For Each j1 In MaxCategsArray
    For Each j2 In TempPbs
     If InArray(j2,j1.ProblemsArray) then j1.ProblemsLoaded = j1.ProblemsLoaded + 1
    Next 
   Next

  Next

  contor = 0
  loopcount = TestDefinition.MaxRandom - (UBound(GeneratedTestArray)+1) + (UBound(IncludePbArray)+1)
  do until contor = loopcount  
    ' Daca "s-au terminat" problemele se iese fortat din bucla do
    On Error Resume Next
    NrAvailPbs = UBound(AvailableProblems) + 1
    If Err.number<>0 then 
      NrAvailPbs = 0
      Err.Clear 
    End If  
    If NrAvailPbs = 0 then Exit Do

    ' Se elimina categoriile "full"
    For Each i In MaxCategsArray
     If i.ProblemsLoaded >= i.ProblemsWanted then AvailableProblems  = ArrayDif(AvailableProblems, i.ProblemsArray)
    Next
    
    ' Se adauga o noua problema
    TempPbs = GetRandomItems(AvailableProblems, 1)
    GeneratedTestArray = AddArrays(GeneratedTestArray, TempPbs)
    AvailableProblems  = ArrayDif(AvailableProblems, TempPbs)
    
    ' Se incrementeaza categoriile in care apare aceasta problema
    For Each i In MaxCategsArray
     If InArray(TempPbs(0),i.ProblemsArray) then i.ProblemsLoaded = i.ProblemsLoaded + 1
    Next
    
    contor = contor + 1
  loop
  
  GenerationDate      = Now()
  GetGeneratedTest    = GeneratedTestArray
 End Function

End Class



' Intoarce un obiect de tip Dictionary cu testele de la un curs
' ce contine pe post de keye ID-ul testului
' iar pe post de item numele ei.
Function GetTotalTstList(cursid, onlypublics, objCon)
 Dim re
 
 Set re = CreateObject("Scripting.Dictionary")
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetTstInfo"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(cursid))
 If not rs.EOF then
   do while not rs.EOF
     If onlypublics then
      If rs.Fields("testpublic").Value then re.Add CStr(rs.Fields("id_test").Value), rs.Fields("numetest").Value
     Else
      re.Add CStr(rs.Fields("id_test").Value), rs.Fields("numetest").Value
     End If
     rs.MoveNext
   loop
 End If  
 rs.Close
 set rs = nothing
 set myCmd = nothing
 
 Set GetTotalTstList = re
End Function


' Intoarce un obiect de tip Dictionary cu testele dintr-o anumita categorie
' ce contine pe post de keye ID-ul testului
' iar pe post de item numele lui.
Function GetCategTstList(categid, onlypublics, objCon)
 Dim re
 
 Set re = CreateObject("Scripting.Dictionary")
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetTstFromCategID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(categid))
 If not rs.EOF then
   do while not rs.EOF
     If onlypublics then
       If rs.Fields("testpublic").Value then re.Add CStr(rs.Fields("id_test").Value), rs.Fields("numetest").Value
     Else  
       re.Add CStr(rs.Fields("id_test").Value), rs.Fields("numetest").Value
     End If
     rs.MoveNext
   loop
 End If  
 rs.Close
 set rs = nothing
 set myCmd = nothing
 
 Set GetCategTstList = re
End Function

' Intoarce numarul de sustineri pentru un test de catre un utilizator
Function GetTestNrSustineriByUser(testid, userid, objCon)
 Dim myCmd, rs
 Dim re
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetTstNrSustineriByUserAndTst"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,Array(CLng(testid),CLng(userid)))
 If not rs.EOF then
  re = rs.Fields("sustineri").Value
 Else
  re = 0
 End If
 rs.Close
 set rs = nothing
 set myCmd = nothing
 GetTestNrSustineriByUser = re
End Function

' Intoarce numarul de sustineri al unui test de catre studenti
Function GetTestNrSustineri(testid, objCon)
 Dim re
 Const SQLSel = "SELECT Count(id_test) AS NrSustineri FROM TBStudentsResults WHERE id_test=@1"
 Set rs = objCon.Execute(Replace(SQLSel, "@1", CStr(testid)))
 If not rs.EOF then
  re = rs.Fields("NrSustineri").Value
 Else
  re = 0
 End If
 GetTestNrSustineri = re
End Function
%>