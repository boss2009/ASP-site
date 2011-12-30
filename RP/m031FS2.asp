<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Report Top Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="140,*" frameborder="No" framespacing="0" rows="*"> 
	<frame name="RPBrowseLeftFrame" scrolling="NO" src="m031F2Hdr.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
	<frame name="RPBrowseRightFrame" scrolling="YES" noresize src="../AC/m001r01menu.asp" >
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>
