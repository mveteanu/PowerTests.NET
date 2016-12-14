<%
' *********************************************************************
' Answers input - server side support. 
' For client side support see PBAnswers.vbs
' Date:  April 10, 2001
' *********************************************************************

' Functions for managing DB answers

' Returns as a String, an HTML snippet that contains the answering area
' This should look like in the moment the question was saved
' If question is not found the "" is returned
Function GetPbAnswersString(pbid, editbtns, objCon)
  Dim re, ServMachete, OptMacheta, ServFinal
  Dim myCmd, rs
  Dim tipr, contor, optstr, optstr2, optar, i

  re = ""
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetPbAnswers"
  myCmd.CommandType = adCmdStoredProc
  Set rs = myCmd.Execute(,CLng(pbid))
  If not rs.EOF then
    OptMacheta  = "<option value='@1'@2>@1</option>"
    If editbtns then
      ServMachete = Array(_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input @3 id='answer@2' name=answer type=radio class='TAnswerButton' title='Bifati raspunsul corect'></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input @3 id='answer@2' name=answer type=checkbox class='TAnswerButton' title='Bifati raspunsurile corecte'></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><select id='answer@2' name=answer rows=1 class='TAnswerComboBox' title='Selectati raspunsul corect'>@3</select>&nbsp;&nbsp;<a href=# onclick='vbscript:EditCombo(answer@2)' unselectable='on' title='Editati lista cu variante de raspuns'>Edit</a></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input id='answer@2' name=answer type=edit value='@3' class='TAnswerEdit' title='Introduceti raspunsul corect'><input id='propanswer@2' type=edit value='@4' style='display:none;'>&nbsp;&nbsp;<a href=# onclick='vbscript:EditProperties(propanswer@2)' unselectable='on' title='Setare proprietati raspuns'>Proprietati</a></span>")
    Else
      ServMachete = Array(_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input @3 type=radio class='TAnswerButton' title='Bifati raspunsul corect'></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input @3 type=checkbox class='TAnswerButton' title='Bifati raspunsurile corecte'></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><select rows=1 class='TAnswerComboBox' title='Selectati raspunsul corect'>@3</select></span>",_
        "<span id='allanswer@2' unselectable='on' style='display:block;height:25px;'><span unselectable='on' style='width:20px;font-weight:bold;'>@1.</span><input READONLY type=edit value='@3' class='TAnswerEdit' title='Introduceti raspunsul corect'></span>")
    End If    
    contor = 1
    do until rs.EOF
     tipr = rs.Fields("tipraspuns").Value
     If tipr = 3 then
       optar = Split(rs.Fields("responsedetails").Value, Chr(3), -1, 1)
       optstr = ""
       For i = 0 to UBound(optar)
        optstr2 = Replace(OptMacheta, "@1", optar(i))
        If i = CInt(rs.Fields("responsecorrect").Value) then
          optstr2 = Replace(optstr2, "@2", " SELECTED ")
        Else
          optstr2 = Replace(optstr2, "@2", "")
        End If  
        optstr = optstr & optstr2
       Next
     End If
     ServFinal = Replace(Replace(ServMachete(tipr-1), "@2", CStr(contor)), "@1", Chr(64+contor))
     select case tipr
       case 1,2 If CBool(LCase(rs.Fields("responsecorrect").Value)) then
                  ServFinal = Replace(ServFinal,"@3", "CHECKED")
                Else
                  ServFinal = Replace(ServFinal,"@3", "")
                End If
       case 3   ServFinal = Replace(ServFinal,"@3", optstr)
       case 4   ServFinal = Replace(Replace(ServFinal,"@4", rs.Fields("responsedetails").Value),"@3", rs.Fields("responsecorrect").Value)
     end select
     re = re & ServFinal & vbCrLf
     contor = contor + 1
     rs.MoveNext
    loop
  GetPbAnswersString = re  
  End If
  rs.Close
  set myCmd = nothing
  set rs = nothing
End Function


' Returns question data as an array of 7 elements
' If question cannot be found in DB, -1 is return on first position
' Returns: return(0) = answer type
'          return(1) = number of answers
'          return(2) = question properties (CSV with separator: #255)
'          return(3) = question text
'          return(4) = course name
'          return(5) = number of categories that contains the question
'          return(6) = number of tests that contains the question
Function GetPBSavedData(pbid, objCon)
 Dim myCmd, rs, lit
 Dim re(7)

 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = objCon
 myCmd.CommandText = "GetPbInfoByID"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute(,CLng(pbid))
 If not rs.EOF then
   re(0) = rs.Fields("tipraspuns").Value
   re(1) = rs.Fields("nransw").Value
   If (re(0) = 1) or (re(0) = 2) then
     lit = "d"
   Else
     If rs.Fields("acceptaraspunspartial").Value then
       lit = "t"
     Else
       lit = "f"
     End If    
   End If  
   re(2) = rs.Fields("numeproblema").Value & Chr(3) & rs.Fields("autorproblema").Value & Chr(3) & lit
   re(3) = rs.Fields("textproblema").Value
   re(4) = rs.Fields("numecurs").Value
   re(5) = rs.Fields("nrcateg").Value
   re(6) = rs.Fields("nrtests").Value
 Else
   re(0) = -1  
 End If
 rs.Close
 set myCmd = nothing
 set rs = nothing
 
 GetPBSavedData = re
End Function


' ==========================================================
' Visual functions
' ==========================================================

' Main sub that opens the control and init some server variables
' Input:
'   iLeft   \ iLeft and iTop - control position
'   iTop    / 
'   iWidth = control width
'   iHeight = control height
'   ImagesDir = path to images forlder used for the toolbar
'   strComponentName = component name
Sub OpenPBAnswersZone(iLeft, iTop, iWidth, iHeight, ImagesDir, strComponentName)
  Const DivA = "<div id='@1' unselectable='on' style='@2 border:inset thin; overflow:auto; padding:8px; FONT-FAMILY:MS Sans Serif;FONT-Size: 8pt;cursor:default;' title='Zona cu raspunsuri'>"
  Dim strStyleDef
  
  HTMLPbImagesDir = ImagesDir
  
  if not(IsNumeric(iLeft) and IsNumeric(iTop)) then
    strStyleDef = "display:inline;position:relative;"
  else
    strStyleDef = "position:absolute;left:"+CStr(iLeft)+"px; top:"+CStr(iTop)+"px;"
  end if   
  strStyleDef = strStyleDef + " width:"+CStr(iWidth)+"px; height:"+CStr(iHeight)+"px;"
  
  Response.Write Replace(Replace(DivA, "@2", strStyleDef), "@1", strComponentName)
  Response.Write VbCrLf
End Sub

' Close the control
Sub ClosePBAnswersZone
  Response.Write "</div>"
End Sub

' Toolbar for inserting and handling images
Sub CreateHTMLProblemToolBar(iLeft, iTop)
%>
<div id="ProblemToolBar" UNSELECTABLE="on"
     class=TForm style="width:265px;height:31px;"
     style="left:<%=CStr(iLeft)%>px;top:<%=CStr(iTop)%>px;">
<span UNSELECTABLE="on"
      class=TLabel style="width:160px;height:13px;"
      style="left:6px;top:8px;">
Question answers/properties:
</span>

<div ID="btnaddanswer" class="TSpeedButton"
     onmouseover="javascript:this.className='TSpeedButtonUp'"
     onmouseout="javascript:this.className='TSpeedButton'"
     onmouseup="javascript:this.className='TSpeedButtonUp'"
     onmousedown="javascript:this.className='TSpeedButtonDown'"
     style="left:166px; top:1px;"
     UNSELECTABLE="on" TITLE="Add possible answer">
<img src="<%=HTMLPbImagesDir%>/bulbon.png" width="23" height="22">
</div>
<div ID="btndelanswer" class="TSpeedButton"
     onmouseover="javascript:this.className='TSpeedButtonUp'"
     onmouseout="javascript:this.className='TSpeedButton'"
     onmouseup="javascript:this.className='TSpeedButtonUp'"
     onmousedown="javascript:this.className='TSpeedButtonDown'"
     style="left:192px; top:1px;"
     UNSELECTABLE="on" TITLE="Delete answer">
<img src="<%=HTMLPbImagesDir%>/bulboff.png" width="23" height="22">
</div>
<img SRC="<%=HTMLPbImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:218px; top:1px;">
<div ID="btnpbproperties" class="TSpeedButton"
     onmouseover="javascript:this.className='TSpeedButtonUp'"
     onmouseout="javascript:this.className='TSpeedButton'"
     onmouseup="javascript:this.className='TSpeedButtonUp'"
     onmousedown="javascript:this.className='TSpeedButtonDown'"
     style="left:224px; top:1px;"
     UNSELECTABLE="on" TITLE="Question properties">
<img src="<%=HTMLPbImagesDir%>/proprietati.png" width="23" height="22">
</div>
</div>
<%
End Sub

' Private statements

Dim HTMLPbImagesDir
%>

