<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Follow Up Notes Frame</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="FollowUpFrameHeader" scrolling="NO"  noresize src="m001F11Hdr.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" >
  <frame name="FollowUpFrameBody" scrolling="yes" src="m001q1101.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>
