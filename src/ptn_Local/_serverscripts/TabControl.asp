<%
' *********************************************************************
' TTabControl server side support
' Date:  March 11, 2001
' Based on an ASP Today article
' *********************************************************************

' ***************************************************
' BEGIN
' PUBLIC
' SECTION
' ***************************************************

' This main sub opens the control
' strLeft          \ strLeft and strTop are used to position the control
' strTop           / if missing the control is positioned relatively
' strWidth         = Control width
' strHeight        = Control height
' strTabNames      = Array with tab header labels
' iTabDefault      = Default opened tab (1 = first tab)
' strComponentName = Component name
Public Sub OpenTabControl(strLeft, strTop, strWidth, strHeight, strTabNames, iTabDefault, strComponentName)
  Dim i
  Dim flgDefault
  Dim strTabName
  Dim strCanvasID, strTabHeaderID, strTabContentID

  TabControl(0) = strTabNames
  TabControl(1) = iTabDefault
  TabControl(2) = strHeight
  TabControl(3) = 0
  
  strCanvasID = "tab"&strComponentName 
  call OpenCanvas( strCanvasID, strLeft, strTop, strWidth, strHeight )

  for i = 0 to UBound(strTabNames)
  
    strTabName = CvtStrToID(strTabNames(i))
    strTabHeaderID = "tabHdr"&strTabName: strTabContentID = "tabCnt"&strTabName  
    
    if( i = ( iTabDefault - 1 ))then flgDefault = true else flgDefault = false
    
    call CreateHeaderTab( strTabHeaderID, strCanvasID, strTabContentID, i+1, strTabNames(i), flgDefault )
  next
End Sub


' Close the control
Public Sub CloseTabControl
  call CloseCanvas
End Sub  


' Open a tab page
Public Sub OpenTabContent
  Dim strId
  Dim strHeight
  Dim strStyleDef
  Dim strTag
  Dim strVisible
  Dim flgDefault

  TabControl(3) = TabControl(3) + 1
  
  strId = "tabCnt"&CvtStrToID(TabControl(0)(TabControl(3)-1) )
  
  strHeight = CStr(TabControl(2)-28)+"px"

  if( TabControl(3) = TabControl(1)) then flgDefault = true else flgDefault = false
  if( flgDefault=true ) then strVisible = "inherit" else strVisible = "hidden"

  strStyleDef = "POSITION: absolute;"
  strStyleDef = strStyleDef + " TOP:28px; LEFT:0px; WIDTH:99%;"
  strStyleDef = strStyleDef + " HEIGHT:"+strHeight+";"
  strStyleDef = strStyleDef + " CURSOR:default; Z-INDEX:2; BACKGROUND-COLOR:buttonface;"
  strStyleDef = strStyleDef + " BORDER:outset thin;"
  strStyleDef = strStyleDef + " VISIBILITY:"+strVisible+"; OVERFLOW: hidden;"
  
  strTag = vbCrLf+ "<DIV"
  strTag = strTag + " id="+strId
  strTag = strTag + " style='"+strStyleDef+"'"
  strTag = strTag + ">"
  
  Response.Write strTag
End Sub


' Close a tab page
Public Sub CloseTabContent
  Response.Write("</DIV>"+vbCrLf)
End Sub


' ***************************************************
' BEGIN
' PRIVATE
' SECTION
' ***************************************************


Dim TabControl(4)


' Makes an ID from the name of a tab
Private Function CvtStrToID(strTabName)
  CvtStrToID = Replace(Replace(Replace(Replace(strTabName,",","_")," ","_"),"(","_"),")","_")
End Function


' Creates DIVs for tab headers
Private Sub CreateHeaderTab( strId, strCanvasId, strCntId, iTab, strTabName, flgDefault )
  Dim iLeft,iTop
  Dim strTag, strStyleDef
  Dim strOnClickEvent
  Dim strCursor, strZindex, strBrdBottom
  Dim strLeft, strTop, strWidth
  
  iLeft = ( iTab-1 )*100: strLeft = CStr( iLeft )+"px"
  strOnClickEvent = " onclick='javascript:tabActivate("+strCanvasId+","+strId+","+strCntId+")'"
  strStyleDef = "POSITION: absolute;"
  
  if( flgDefault=true ) then
    strTop = "4px": strWidth = "108px": strZindex="3"
    strCursor = "default"
    strBrdBottom = "buttonface solid 1px"
    strStyleDef = strStyleDef+ " CURSOR:"+strCursor+"; Z-INDEX:"+strZindex+"; BORDER-BOTTOM:"+strBrdBottom+";"
  else
    strTop = "8px": strWidth = "100px": strZindex = "1"
    strCursor = "hand"
    strStyleDef = strStyleDef+ " CURSOR:"+strCursor+"; Z-INDEX:"+strZindex+";"
  end if
  
  strStyleDef = strStyleDef+" LEFT:"+strLeft+"; TOP:"+strTop+"; WIDTH:"+strWidth+";"
  strStyleDef = strStyleDef+" HEIGHT:28px; BACKGROUND-COLOR:buttonface;"
  strStyleDef = strStyleDef+" FONT-FAMILY:verdana; FONT-SIZE:8pt; FONT-WEIGHT:bolder; "
  strStyleDef = strStyleDef+" PADDING-BOTTOM:1px; PADDING-TOP:1px; "
  strStyleDef = strStyleDef+" BORDER-TOP:outset thin; BORDER-LEFT:outset thin; BORDER-RIGHT:outset thin;"
  strStyleDef = strStyleDef+" TEXT-ALIGN:center;"
  
  strTag = "<DIV"
  strTag = strTag + " id="+strId
  strTag = strTag + " style='"+strStyleDef+"'"
  strTag = strTag + strOnClickEvent
  strTag = strTag + ">"
  strTag = strTag + strTabName
  strTag = strTag + "</DIV>" + vbCrLf

  Response.Write strTag
End Sub


' Creates the main DIV where headers and pages will stay
Private Sub OpenCanvas( strCanvasId, strLeft, strTop, strWidth, strHeight )
  Dim strStyleDef
  Dim strTag
  
  if not(IsNumeric(strLeft) and IsNumeric(strTop)) then
     strStyleDef = "DISPLAY:inline;position:relative;"
  else
     strStyleDef = "position:absolute;LEFT:"+CStr(strLeft)+"px; TOP:"+CStr(strTop)+"px;"
  end if   
  strStyleDef = strStyleDef + " WIDTH:"+CStr(strWidth)+"px; HEIGHT:"+CStr(strHeight)+"px;"
  strStyleDef = strStyleDef +"Z-INDEX:0;"
  
  strTag = "<DIV"
  strTag = strTag + " id="+strCanvasId
  strTag = strTag + " name="+strCanvasId
  strTag = strTag + " style='"+strStyleDef+"'"
  strTag = strTag + ">" + vbCrLF+vbCrLF
  
  Response.Write strTag
End Sub


' Close main DIV
Private Sub CloseCanvas
  Response.Write (vbCrLf+"</DIV>"+vbCrLf)
End Sub

%>
