<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Client Top Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="140,*" frameborder="No" framespacing="0" rows="*"> 
	<frame name="AdultClientBrowseLeftFrame" scrolling="NO" src="m001F2Hdr.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
	<frame name="AdultClientBrowseRightFrame" scrolling="YES" noresize src="m001s0101.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" >
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>
