<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/HTMLEditControl.asp" -->
<!-- #include file="../_serverscripts/HTMLPbControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim cn, tar

Dim ServTipRaspuns
Dim ServNumarRasp
Dim ServPbProperties
Dim ServPbText
Dim ServPbAnsw
Dim ServPbID

If Request.QueryString.Count = 0 then
 ServTipRaspuns    = 0
 ServNumarRasp     = 0
 ServPbProperties  = Chr(3) & Chr(3) & "f"
 ServPbText        = "Type here the question..."
 ServPbAnsw        = ""
 ServPbID          = ""
Else
 set cn = Server.CreateObject("ADODB.Connection") 
 cn.Open Application("DSN")
 tar = GetPBSavedData(Request.QueryString("PBId"), cn)
 If tar(0) = -1 then 
   ServTipRaspuns    = 0
   ServNumarRasp     = 0
   ServPbProperties  = Chr(3) & Chr(3) & "f"
   ServPbText        = "<font color=red>Question not found!</font>"
 Else
   ServTipRaspuns    = tar(0)
   ServNumarRasp     = tar(1)
   ServPbProperties  = tar(2)
   ServPbText        = tar(3)
 End If
 ServPbAnsw = GetPbAnswersString(Request.QueryString("PBId"), true, cn)
 ServPbID   = "value=" & Request.QueryString("PBId")
 cn.Close
 set cn = nothing
End If
%>
<html>
<head>
  <title>Problem Editor</title>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>

<body unselectable="on" style="behavior:url('../_clientscripts/application.htc');">

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Please wait...
</td></tr></table>
</div>

<div id="Form1" style="overflow:hidden;visibility:hidden;" unselectable="on"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<%
OpenPBAnswersZone 7, 375, 744, 100, "../images", "AnsDIV"
Response.Write ServPbAnsw
ClosePBAnswersZone

CreateHTMLProblemToolBar 6, 3

CreateChosePictureForm "pbupload_ser.asp", "target='FormReturn'", _
                       "<input type=hidden id='idopenpb' name='idopenpb' "& ServPbID &">" &_
                       "<input type=hidden id='delimglist' name='delimglist'>" &_
                       "<input type=hidden id='uploadlist' name='uploadlist'>" &_
                       "<input type=hidden id='pbtext' name='pbtext'>" &_
                       "<input type=hidden id='pbprop' name='pbprop'>" &_
                       "<input type=hidden id='pbansw' name='pbansw'>"
OpenHTMLEditControl 6,70,744,300, "../_clientscripts/HTMLEditor.htc", "../images", "MyHTMLEdit"
OpenPage 705,275,""
Response.Write ServPbText
ClosePage
CloseHTMLEditControl
CreateHTMLPicturesToolBar 270,3
CreateHTMLZoomToolBar 609, 3
CreateHTMLEditToolBar 6,33
%>

<input id="ButtonSave" type=button value="Save" title="Save question"
       class=TButton style="width:75px;height:25px;"
       style="left:297px;top:485px;">
<input id="ButtonClose" type=button value="Close" title="Close form"
       class=TButton style="width:75px;height:25px;"
       style="left:385px;top:485px;">
</div>

<div id="Form1Hidden" style="display:none;">
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language="vbscript" src="../_clientscripts/PBAnswers.vbs"></script>
<script language="vbscript" src="../_clientscripts/HTMLEditorUtils.vbs"></script>

<script language="vbscript">
Dim ServerImagesAtLoadTime, ServerTextAtLoadTime
Dim ServerPropAtLoadTime, ServerNrAnsAtLoadTime, ServerAnswAtLoadTime
Dim NumarRasp
Dim TipRaspuns
Dim PbProperties

Sub Window_onload
 Form1.style.visibility = "visible"
 WaitforForm.style.visibility = "hidden"

 NumarRasp    = <%=ServNumarRasp%>
 TipRaspuns   = <%=ServTipRaspuns%>
 PbProperties = "<%=ServPbProperties%>"
 ServerImagesAtLoadTime = GetEditPageServerImages(MyHTMLEdit_TextBox1)
 ServerTextAtLoadTime   = MyHTMLEdit_TextBox1.innerHTML
 ServerPropAtLoadTime   = PbProperties
 ServerNrAnsAtLoadTime  = NumarRasp
 ServerAnswAtLoadTime   = GetAnswers(AnsDIV, NumarRasp, TipRaspuns)
 MyHTMLEdit_TextBox1.ContentEditable = true
End Sub


Sub btnaddanswer_onclick
  AddAnswer AnsDIV, NumarRasp, TipRaspuns
End Sub

 
Sub btndelanswer_onclick
  RemoveAnswer AnsDIV, NumarRasp
End Sub


Sub btnpbproperties_onclick
  Dim re, spr
  
  If NumarRasp < 1 then 
    Msgbox "First you need to define the possible answers for this question" & vbCrLf & "and then you will be able to set question properties.", vbOKOnly+vbExclamation,"Info"
    Exit Sub
  End If
  
  spr = Split(PbProperties, Chr(3), -1, 1)
  If (TipRaspuns = 1) or (TipRaspuns = 2) then
    PbProperties = spr(0) & Chr(3) & spr(1) & Chr(3) & "d"
  ElseIf spr(2)="d" then 
    PbProperties = spr(0) & Chr(3) & spr(1) & Chr(3) & "f"
  End If
  re = ShowModalDialog("pbproprpb.asp", PbProperties, "dialogWidth=330px;dialogHeight=240px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If VarType(re) = vbString then PbProperties = re
End Sub

' =====================================================

' Inchide problema neconditionat
Sub CloseProblemNeconditionat
  MyHTMLEdit_TextBox1.ContentEditable = false
  Window.returnValue = 0
  Window.close 
End Sub

' Inchide problema salvata
Sub CloseSavedProblem
  MyHTMLEdit_TextBox1.ContentEditable = false
  Window.returnValue = 1
  window.close 
End Sub

' Evenimentul apare la apasarea butonului Close
Sub ButtonClose_onclick
  Dim wasChanged
  
  If (ServerTextAtLoadTime = MyHTMLEdit_TextBox1.innerHTML) and (ServerPropAtLoadTime = PbProperties) and (ServerNrAnsAtLoadTime = NumarRasp) and (ServerAnswAtLoadTime = GetAnswers(AnsDIV, NumarRasp, TipRaspuns)) then
    wasChanged = false
  Else
    wasChanged = true
  End If    

  If wasChanged then 
   Select case msgbox("Question was changed. Do you want to save it?",vbYesNoCancel + vbInformation, "Warning")
     case vbYes ButtonSave_onclick
     case vbNo  CloseProblemNeconditionat
   End Select
  Else
   CloseProblemNeconditionat 
  End If
End Sub


' Evenimentul apare la apasarea butonului Save
Sub ButtonSave_onclick
  Dim deletimg, localimg
  Dim deletstring, localstring
  Dim pbanswers

  pbanswers = GetAnswers(AnsDIV, NumarRasp, TipRaspuns)
  If pbanswers="" then
    msgbox "You need to define possible answers"& vbCrLf &"and then specify the correct one(s).", vbOkOnly+vbExclamation, "Info"
    Exit Sub
  End If     

  deletimg = ArrayDif(ServerImagesAtLoadTime, GetEditPageServerImages(MyHTMLEdit_TextBox1))
  localimg = GetEditPageLocalImages(MyHTMLEdit_TextBox1)

  If Join(deletimg)="" then
   deletstring = ""
  else
   deletstring = FileNamesArrayToIDCSV(deletimg, "FileID")
  end if

  CleanFilesForm ChosePictureFormular,localimg
  localstring = GetUploadFields(ChosePictureFormular)

  ChosePictureFormular.delimglist.value = deletstring
  ChosePictureFormular.uploadlist.value = localstring
  ChosePictureFormular.pbtext.value = MyHTMLEdit_TextBox1.innerHTML
  ChosePictureFormular.pbprop.value = PbProperties
  ChosePictureFormular.pbansw.value = pbanswers
  ChosePictureFormular.Submit
End Sub
</script>

</body>
</html>

