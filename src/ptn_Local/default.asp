<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
  <title>PowerTests .NET</title>
  <link rel="stylesheet" type="text/css" href="css/ptn.css">
  <link rel="SHORTCUT ICON" href="images/ptn.ico">
  <meta name="Description" content="PowerTests .NET">
  <meta name="Keywords" content="powertests, teste, educatie">
</head>

<body bgcolor="white" class="white">

<table width="100%" height="100%" border="0">
<tr><td align="center" valign="middle">
  <table width="600" height="240" border="1" cellspacing="0" cellpadding="10" bordercolor="Black" style="border: 1px solid Black;">
    <tr><td align="left" valign="top"><font face="Verdana,Arial" size="-1">
	  <table border="0" width="100%">
	  <tr>
	  <td align="center" valign="middle" width="110"><img src="images/ptnlogo.png" width="105" height="37" border="0" align="middle" alt="PowerTests .NET"></td>
	  <td align="center" valign="middle"><font face="Verdana,Arial" size="-1"><b>PowerTests .NET<br>Revolutionary Assessment Solution</b></font></td>
	  </tr>
	  </table>
	  <p align="left">
	  Welcome to <i>PowerTests .NET</i>. If you are a student you can verify your knowledge by 
	  enrolling at your preferred course now. If you are a professor, you can use the modern and easy 
	  to use included tools to verify your student’s knowledge .
	  </p>
		<p align="left">
		<font size="-2"><u>Note:</u> To use the application you need an account. If you don’t have one 
		you can sign-up now in the next screen.</font>
		</p>
	  <center>
	  <form name=launchform>
	   <input CHECKED type="checkbox" id="fullscreen">Run in fullscreen
	   <br>
	   <input type="button" value="Start application" title="Launch PowerTests .NET !" onclick="EnterSite()">
	  </form>
	  </center>
	</td></tr>
  </table>
</td></tr>
</table>

<script LANGUAGE="JavaScript">
<!--
// Intoarce true daca browserul este Internet Explorer si
// versiunea este cel putin minvers
function IsIE(minvers)
{
 var bIsIE;
 sAgent = navigator.userAgent;
 if((i1 = sAgent.indexOf("MSIE"))!=-1)
  {
	i2 = sAgent.indexOf(";", i1);
	sAgVer = parseFloat(sAgent.slice(i1+5, i2));
	if (sAgVer=="NaN") bIsIE = false;
	else 
	 {
	   if(sAgVer>=minvers) bIsIE = true;
	   else bIsIE = false;
	 }
  }
 else bIsIE = false;
 return(bIsIE); 
}

function EnterSite()
 {	
  if (IsIE(5.5))
   { 
    if(launchform.fullscreen.checked) bIsLaunched = window.open("login/default.asp",null,"fullscreen=yes, toolbar=no, menubar=no, location=no, status=no");
    else bIsLaunched = window.open("login/default.asp",null,"fullscreen=no, width=800, height=600, toolbar=no, menubar=no, location=no, status=no, resizable=yes");
   } 
  else
   alert("PowerTest .NET ruleaza doar in\nMicrosoft Internet Explorer 5.5+\n\nVa multumim pentru intelegere!")
 }
//-->
</script>


</body>
</html>
