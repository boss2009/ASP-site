<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Error Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Error</h5>
<hr>
<i>System has detected a duplicate Inventory ID.  Inventory has not been created.  Please go back and enter another ID.</i><br></br>
<input type="button" value="Back" onclick="history.go(-1);" class="btnstyle" >
</body>
</html>
