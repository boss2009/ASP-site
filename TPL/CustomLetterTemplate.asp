<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/CustomLetterHeader.inc" -->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Custom Letter Template</title>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/i/letterhead.gif" width="450" height="80"></p>
<%=Custom_Letter_Content%>
</body>
</html>