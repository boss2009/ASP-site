<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Documentation Eligibility Frame Set</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="DocumentationEligibilityFrameHeader" scrolling="NO"  noresize src="m001F06Hdr.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" >
	<frame name="DocumentationEligibilityFrameBody" scrolling="yes" src="m001q0601.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.</body>
</noframes> 
</html>
