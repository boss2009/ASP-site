<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body onload="setTimeout('self.close()',3000); window.opener.location.reload();">
<b>The record has been
<%
if (String(Request("action")) == "Abort" ) {
	Response.Write("removed previously, action aborted");
} else {
	if (String(Request("action")) != "undefined" ) {
		Response.Write(" "+Request.QueryString("action")+" ");
	} else { 
		Response.Write(" Update ?...");
	}
	Response.Write(" successfully.");
}
%>
</b>
<br><br>
This window will close in 3 seconds<br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
