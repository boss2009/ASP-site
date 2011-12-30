<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>System Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function Init(){
		top.HeaderFrame.location.reload();	
		window.location.href='<%=Request.QueryString("page")%>?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intReferral_id=<%=Request.QueryString("intReferral_id")%>';
		//history.back();	
	}
	</script>
</head>
<body onload="setTimeout('Init();',1000);">
Record successfuly updated.
</body>
</html>
