<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Loan Top Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="140,*" frameborder="No" framespacing="0" rows="*">
	<frame name="LoanBrowseLeftFrame" scrolling="NO" src="m008F2Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
	<frame name="LoanBrowseRightFrame" scrolling="YES" noresize src="m008s0101.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" >
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
