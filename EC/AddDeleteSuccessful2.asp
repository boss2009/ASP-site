<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>System Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function Init(){
		setTimeout("window.location.href='m007q01lw.asp'",2000);
	}
	</script>
</head>
<body onload="Init();">
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
This form will close in 2 seconds.<br><br>
</body>
</html>
