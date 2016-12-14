<%
' *********************************************************************
' The server side implementation of THTMLEditor control
' Date:  April 02, 2001
' *********************************************************************

' ***************************************************
' BEGIN
' PUBLIC
' SECTION
' ***************************************************

' This is the main sub that opens the control and init some server variables
' Input:
'   iLeft   \ iLeft and iTop - position of the control
'   iTop    / if these are missing the control is positioned relativelly
'   iWidth = control width
'   iHeight = control height
'   HTCName = path to .HTC script that handles the client side events
'   ImagesDir = path to images folder used on editor toolbar
'   strComponentName = component name
'Option Explicit

Public Sub OpenHTMLEditControl(iLeft, iTop, iWidth, iHeight, HTCName, ImagesDir, strComponentName)
  Dim strStyleDef
  Dim strDiv1, strDiv2
 
  strDiv1  = "<div ID='@1' class='@2' style='@3'>" & vbCrLf
  strDiv2  = "<div ID='@1' class='@2'>" & vbCrLf
  HTMLEditWidth      = iWidth
  HTMLEditHeight     = iHeight
  HTMLCompName       = strComponentName
  HTMLHTCName        = HTCName
  HTMLImagesDir      = ImagesDir
  HTMLEditCurentPage = 0
     
  if not(IsNumeric(iLeft) and IsNumeric(iTop)) then
    strStyleDef = "display:inline;position:relative;"
  else
    strStyleDef = "position:absolute;left:"+CStr(iLeft)+"px; top:"+CStr(iTop)+"px;"
  end if   
  strStyleDef = strStyleDef + " width:"+CStr(iWidth)+"px; height:"+CStr(iHeight)+"px;"

  
  Response.Write Replace(Replace(Replace(strDiv1,"@3",strStyleDef),"@2","THTMLEditMainContainer"),"@1", strComponentName) &_
                 Replace(Replace(strDiv2,"@2","THTMLEditZoomContainer"),"@1", strComponentName & "_ZoomContainer")
End Sub


' Close the control
Public Sub CloseHTMLEditControl
 Response.Write "</div>" & vbCrLf &_
                "</div>"
End Sub


' Starts a new editing page inside the control
' strWaterMark is the path to a watermark image that can be displayed
' on the page. If no watermark is required then strWaterMark = ""
Public Sub OpenPage(iWidth, iHeight, strWaterMark)
 Dim strDiv1, strDiv2, strDiv3, strDiv4
 
 HTMLEditCurentPage = HTMLEditCurentPage + 1
 strDiv1 = "<div ID='@0_PageBorder@1' class='THTMLEditPageBorder' style='width:@2px; height:@3px;'>" & vbCrLf
 strDiv2 = "<div ID='@0_Page@1' class='THTMLEditPage' style='width:@2px; height:@3px;'>" & vbCrLf
 strDiv3 = "<span CONTENTEDITABLE=false ID='@0_TextBox@1' class='THTMLEditTextBox' style='width:@2px; height:@3px;'>"

 strDiv4 = "<div STYLE='position:absolute;width:@1px;height:@2px;z-index:-1;overflow:hidden;'>" & vbCrLf &_
           "<table width=100% height=100% border=0 cellpadding=0 cellspacing=0>" & vbCrLf  &_
           "<tr><td align=center valign=center><img SRC='@3'></td></tr></table></div>" & vbCrLf 

 strDiv1 = Replace(Replace(Replace(Replace(strDiv1,"@0",HTMLCompName), "@3", CStr(iHeight)), "@2", CStr(iWidth)), "@1", CStr(HTMLEditCurentPage))
 strDiv2 = Replace(Replace(Replace(Replace(strDiv2,"@0",HTMLCompName), "@3", CStr(iHeight)), "@2", CStr(iWidth)), "@1", CStr(HTMLEditCurentPage))
 strDiv3 = Replace(Replace(Replace(Replace(strDiv3,"@0",HTMLCompName), "@3", CStr(iHeight-40)), "@2", CStr(iWidth-40)), "@1", CStr(HTMLEditCurentPage))
 strDiv4 = Replace(Replace(Replace(strDiv4, "@2", CStr(iHeight)), "@1", CStr(iWidth)), "@3", strWaterMark)

 Response.Write strDiv1
 Response.Write strDiv2
 if strWaterMark<>"" then Response.Write strDiv4
 Response.Write strDiv3
End Sub


' Closes a page
Public Sub ClosePage
 With Response
  .Write "</span>" & vbCrLf
  .Write "</div>" & vbCrLf
  .Write "</div>" & vbCrLf
 End With
End Sub


' Private statements

Dim HTMLEditWidth
Dim HTMLEditHeight
Dim HTMLEditCurentPage
Dim HTMLCompName
Dim HTMLHTCName
Dim HTMLImagesDir


' Toolbars declarations
' In the consumer .asp page these should be invoked
' after the main call that opens the control

' Defines the formatting toolbar
Sub CreateHTMLEditToolBar(iLeft, iTop)
%>
<div id="EditToolBar" UNSELECTABLE="on"
     class=TForm style="width:744px;height:31px;"
     style="left:<%=CStr(iLeft)%>px;top:<%=CStr(iTop)%>px;">

<div ID="btnbold" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:352px; top:1px;"
     UNSELECTABLE="on" TITLE="Bold">
<img src="<%=HTMLImagesDir%>/editbold.png" width="23" height="22">
</div>

<div ID="btnitalic" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:378px; top:1px;"
     UNSELECTABLE="on" TITLE="Italic">
<img src="<%=HTMLImagesDir%>/edititalic.png" width="23" height="22">
</div>

<div ID="btnunderline" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:404px; top:1px;"
     UNSELECTABLE="on" TITLE="Underline">
<img src="<%=HTMLImagesDir%>/editunderline.png" width="23" height="22">
</div>

<div ID="btnstrike" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:430px; top:1px;"
     UNSELECTABLE="on" TITLE="Strike">
<img src="<%=HTMLImagesDir%>/editstrike.png" width="23" height="22">
</div>

<img SRC="<%=HTMLImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:456px; top:1px;">

<div ID="btnsuperscript" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:462px; top:1px;"
     UNSELECTABLE="on" TITLE="Superscript">
<img src="<%=HTMLImagesDir%>/editsuperscript.png" width="23" height="22">
</div>

<div ID="btnsubscript" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:488px; top:1px;"
     UNSELECTABLE="on" TITLE="Subscript">
<img src="<%=HTMLImagesDir%>/editsubscript.png" width="23" height="22">
</div>

<img SRC="<%=HTMLImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:514px; top:1px;">

<div ID="btnalignleft" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:520px; top:1px;"
     UNSELECTABLE="on" TITLE="Align Left">
<img src="<%=HTMLImagesDir%>/editleftalign.png" width="23" height="22">
</div>

<div ID="btnaligncenter" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:546px; top:1px;"
     UNSELECTABLE="on" TITLE="Center">
<img src="<%=HTMLImagesDir%>/editcenteralign.png" width="23" height="22">
</div>

<div ID="btnalignright" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:572px; top:1px;"
     UNSELECTABLE="on" TITLE="Align Right">
<img src="<%=HTMLImagesDir%>/editrightalign.png" width="23" height="22">
</div>

<div ID="btnalignvertical" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:598px; top:1px;"
     UNSELECTABLE="on" TITLE="Vertical Text">
<img src="<%=HTMLImagesDir%>/editvertical.png" width="23" height="22">
</div>

<img SRC="<%=HTMLImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:624px; top:1px;">

<div ID="btnorderedlist" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:630px; top:1px;"
     UNSELECTABLE="on" TITLE="Ordered List">
<img src="<%=HTMLImagesDir%>/editnumberlist.png" width="23" height="22">
</div>

<div ID="btnunorderedlist" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:656px; top:1px;"
     UNSELECTABLE="on" TITLE="Unordered List">
<img src="<%=HTMLImagesDir%>/editbulletlist.png" width="23" height="22">
</div>

<div ID="btnoutdent" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:682px; top:1px;"
     UNSELECTABLE="on" TITLE="Outdent">
<img src="<%=HTMLImagesDir%>/editoutdent.png" width="23" height="22">
</div>

<div ID="btnindent" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:708px; top:1px;"
     UNSELECTABLE="on" TITLE="Indent">
<img src="<%=HTMLImagesDir%>/editindent.png" width="23" height="22">
</div>

<select id="combofontface" UNSELECTABLE="on" title="Font"
        class=TComboBox style="width:105px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:8px;top:3px;">
<option value="Times New Roman" SELECTED>Times New Roman</option>
<option value="Arial">Arial</option>
<option value="Verdana">Verdana</option>
<option value="Courier New">Courier New</option>
<option value="Symbol">Symbol</option>
<option value="Webdings">Webdings</option>
</select>


<select id="combofontsize" UNSELECTABLE="on" title="Font Size"
        class=TComboBox style="width:57px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:114px;top:3px;">
<option value="1">8</option>
<option value="2">10</option>
<option value="3" SELECTED>12</option>
<option value="4">14</option>
<option value="5">16</option>
<option value="6">18</option>
<option value="7">20</option>
</select>


<select id="combocolor" UNSELECTABLE="on"  title="Text Color"
        class=TComboBox style="width:80px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:180px;top:3px;">
<option VALUE="#000000" STYLE="color:black" SELECTED>Black</option>
<option VALUE="#0000cd" STYLE="color:mediumblue">Blue</option>
<option VALUE="#228b22" STYLE="color:forestgreen">Green</option>
<option VALUE="#ffff00" STYLE="color:yellow">Yellow</option>
<option VALUE="#ff8c00" STYLE="color:darkorange">Orange</option>
<option VALUE="#dc143c" STYLE="color:crimson">Red</option>
<option VALUE="#9400d3" STYLE="color:darkviolet">Purple</option>
<option VALUE="#808080" STYLE="color:gray">Gray</option>
<option VALUE="#ffffff" STYLE="color:white">White</option>
</select>


<select id="combobgcolor" UNSELECTABLE="on"  title="Text BackgroundColor"
        class=TComboBox style="width:80px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:261px;top:3px;">
<option VALUE="#ffffff" STYLE="background-color:white; color:black" SELECTED>White</option>
<option VALUE="#000000" STYLE="background-color:black;">Black</option>
<option VALUE="#1e90ff" STYLE="background-color:dodgerblue">Blue</option>
<option VALUE="#32cd32" STYLE="background-color:limegreen">Green</option>
<option VALUE="#ffff00" STYLE="background-color:yellow">Yellow</option>
<option VALUE="#ffa500" STYLE="background-color:orange">Orange</option>
<option VALUE="#ff0000" STYLE="background-color:red">Red</option>
<option VALUE="#9370db" STYLE="background-color:mediumpurple">Purple</option>
<option VALUE="#c0c0c0" STYLE="background-color:#c0c0c0;">Gray</option>
</select>

</div>     
<%
End Sub  ' CreateHTMLEditToolBar


' Defines images toolbar
Sub CreateHTMLPicturesToolBar(iLeft, iTop)
%>
<div id="ImagesToolBar" UNSELECTABLE="on"
     class=TForm style="width:340px;height:31px;"
     style="left:<%=CStr(iLeft)%>px;top:<%=CStr(iTop)%>px;">
<span UNSELECTABLE="on"
      class=TLabel style="width:44px;height:13px;"
      style="left:6px;top:8px;">
Images:
</span>

<div ID="btninsertimage" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:50px; top:1px;"
     UNSELECTABLE="on" TITLE="Insert picture">
<img src="<%=HTMLImagesDir%>/editinsertpicture.png" width="23" height="22">
</div>

<img SRC="<%=HTMLImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:76px; top:1px;">

<div ID="btnpictureleftalign" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:82px; top:1px;"
     UNSELECTABLE="on" TITLE="Picture Left Align">
<img src="<%=HTMLImagesDir%>/editpictureleftalign.png" width="23" height="22">
</div>

<div ID="btnpicturenoalign" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:108px; top:1px;"
     UNSELECTABLE="on" TITLE="Picture No Align">
<img src="<%=HTMLImagesDir%>/editpicturenoalign.png" width="23" height="22">
</div>

<div ID="btnpicturerightalign" class="TSpeedButton"
     style="behavior:url('<%=HTMLHTCName%>');"
     style="left:134px; top:1px;"
     UNSELECTABLE="on" TITLE="Picture Right Align">
<img src="<%=HTMLImagesDir%>/editpicturerightalign.png" width="23" height="22">
</div>

<img SRC="<%=HTMLImagesDir%>/editdivider.png" UNSELECTABLE="on" 
     ALIGN="absmiddle" HSPACE="2" width="2" height="25"
     style="position:absolute; left:160px; top:1px;">

<span UNSELECTABLE="on"
      class=TLabel style="width:40px;height:13px;"
      style="left:166px;top:8px;">
Border:
</span>

<select id="comboborderwidth" UNSELECTABLE="on" title="Picture Border Width"
        class=TComboBox style="width:40px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:207px;top:3px;">
<option value="0" SELECTED>0</option>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
</select>

<select id="combobordercolor" UNSELECTABLE="on" title="Picture Border Color"
        class=TComboBox style="width:80px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:248px;top:3px;">
<option VALUE="#000000" STYLE="color:black" SELECTED>Black</option>
<option VALUE="#0000cd" STYLE="color:mediumblue">Blue</option>
<option VALUE="#228b22" STYLE="color:forestgreen">Green</option>
<option VALUE="#ffff00" STYLE="color:yellow">Yellow</option>
<option VALUE="#ff8c00" STYLE="color:darkorange">Orange</option>
<option VALUE="#dc143c" STYLE="color:crimson">Red</option>
<option VALUE="#9400d3" STYLE="color:darkviolet">Purple</option>
<option VALUE="#808080" STYLE="color:gray">Gray</option>
<option VALUE="#ffffff" STYLE="color:white">White</option>
</select>

</div>
<%
End Sub ' CreateHTMLPicturesToolBar


' Defines Zoom toolbar
Sub CreateHTMLZoomToolBar(iLeft, iTop)
%>
<div id="ZoomToolBar" UNSELECTABLE="on"
     class=TForm style="width:141px;height:31px;"
     style="left:<%=CStr(iLeft)%>px;top:<%=CStr(iTop)%>px;">
<span UNSELECTABLE="on"
      class=TLabel style="width:40px;height:13px;"
      style="left:6px;top:8px;">
Zoom:
</span>
<select id="combozoom" UNSELECTABLE="on" title="Zoom percent"
        class=TComboBox style="width:80px;"
        style="behavior:url('<%=HTMLHTCName%>');"
        style="left:50px;top:3px;">
<option value="10%">10%</option>
<option value="25%">25%</option>
<option value="50%">50%</option>
<option value="75%">75%</option>
<option value="80%">80%</option>
<option value="90%">90%</option>
<option value="100%" SELECTED>100%</option>
<option value="120%">120%</option>
<option value="150%">150%</option>
<option value="200%">200%</option>
<option value="500%">500%</option>
</select>
</div>

<script language=vbscript>
Sub combozoom_onchange
 <%=HTMLCompName%>_ZoomContainer.style.zoom = combozoom(combozoom.selectedIndex).value
End Sub
</script>
<%
End Sub  ' CreateHTMLZoomToolBar


' Created the miniform for adding an image
Sub CreateChosePictureForm(FormAction, AditionalFormAttrs, AditionalFormItems)
%>
<div   id="ChosePictureForm" class="TForm" unselectable="on"
       style="visibility:hidden;z-index:10;"
       style="width:300px;height:140px;" 
       style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:200px;">
<span  class=TLabel style="width:90px;height:13px;"
       style="left:20px;top:25px;">Alegeti imaginea:</span>
<input id="ChosePictureFormButOK" type=button value="Accepta" title="Foloseste imaginea selectata"
       class=TButton style="width:75px;height:25px;"
       style="left:70px;top:90px;"
       onclick="javascript:this.parentElement.style.visibility='hidden'">
<input type=button value="Renunta" title="Renunta la imagine"
       class=TButton style="width:75px;height:25px;"
       style="left:155px;top:90px;"
       onclick="javascript:this.parentElement.style.visibility='hidden'" id=button1 name=button1>
<form  id="ChosePictureFormular" method="post" encType="multipart/form-data" action="<%=FormAction%>" <%=AditionalFormAttrs%>>
<%=AditionalFormItems%>
</form>
</div>
<%
End Sub
%>
