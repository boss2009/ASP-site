<%@language="JavaScript"%>
<html>
<head>
	<title>System Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function Init(){
		top.window.location.href="m005FS3.asp?insBundle_id=<%=Request.QueryString("insBundle_id")%>";
	}
	</script>
</head>
<body onload="Init();">
</body>
</html>
