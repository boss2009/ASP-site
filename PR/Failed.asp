<!--------------------------------------------------------------------------
* File Name: Failed.asp
* Title: System Message
* Main SP: 
* Description: When an error occurs in the transacion object, this page is
* displayed.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>System Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body onload="setTimeout('self.close()',3000); window.opener.location.reload();">
<b>An error occured during insertion.  Transaction has been aborted and all inventories have been uncreated.</b>
<br><br>
This window will close in 3 seconds<br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>