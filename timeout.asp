<!--------------------------------------------------------------------------
* File Name: timeout.asp
* Title: Session Timeout
* Main SP: 
* Description: This page is shown when user session times out.
* Author: D. T. Chan
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="inc/ASPUtility.inc" -->
<html>
<head>
	<title>Session Timeout</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	function Relogin(){
		location.href="/../asprelogin.asp?<%=Request.QueryString%>";
	}
	</script>
</head>
<body>
<h5>Session Timeout</h5>
<i>Either session has timed out or user authentification failed.<br>
Close window and login again.<br>
<br>
If you have unsaved work, right click on window and select <b>Back</b><br>
to copy your input.</i>
<br>
<br>
<hr>
<input type="button" value="Close" onClick="top.window.close();" class="btnstyle">
&nbsp;&nbsp;
<input type="button" value="Relogin" onClick="Relogin();" class="btnstyle">
<!--
FilterQuotes(Request.QueryString) seems to return empty string.
<input type="button" value="Relogin" onClick="javascript: location.href='/../asplogin.asp?<%=FilterQuotes(Request.QueryString)%>';" class="btnstyle">
-->
</body>
</html>