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
		window.location.href='<%=Request.QueryString("page")%>?ClassID=<%=Request.QueryString("ClassID")%>';
	}
	</script>
</head>
<body onload="Init();">
Record successfuly updated.
</body>
</html>
