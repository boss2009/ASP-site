<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Documentation Eligibility Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="0" cellspacing="1">
	<tr> 
		<td nowrap><a href="m001q0601.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">Disability Doc</a>|</td>
		<td nowrap><a href="m001q0602.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">Education Doc</a>|</td>
		<td nowrap><a href="m001q0603.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">Grant Eligibility</a>|</td>
		<td nowrap><a href="m001q0604.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">Waiver</a>|</td>
		<td nowrap><a href="m001q0605.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">External Agency</a>|</td>
		<td nowrap><a href="m001q0606.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="DocumentationEligibilityFrameBody">Loan Own Form</a>|</td>
		<td nowrap><a href="m001q0607.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="_blank">Doc Summary</a></td>		
	</tr>
</table>
</body>
</html>