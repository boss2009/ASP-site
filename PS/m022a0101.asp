<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>New Temp Student</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
</head>
<body>
<h5>New Temp Student:</h5>
<hr>
Is this student an existing client?<br><br>
<input type="button" value="Yes" onClick="window.location.href='m022a0103.asp';" tabindex="1" class="btnstyle">&nbsp;&nbsp;
<input type="button" value="No" onclick="window.location.href='m022a0102.asp?IsNew=Yes';" tabindex="2" class="btnstyle">&nbsp;&nbsp;
<input type="button" value="Cancel" onClick="window.close();" tabindex="3" class="btnstyle">
</body>
</html>