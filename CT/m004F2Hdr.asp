<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<html>
<head>
	<title>Contact Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="javascript">
	function openWindow(page){
		if (page!='nothing') loadingstatus=window.open(page, "", "width=240,height=100,scrollbars=0,left=300,top=200,status=0");
	}
	</script>
</head>
<body>
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/tn_CONTACT_02.jpg" ALT="Return to Main Menu." width="80" height="60" border=0></a></td>
	</tr>
	<tr>
		<td height="18px" align="center" class="MenuItem" width="120"><a href="m004d01.asp" target="ContactBodyFrame">Desktop</a></td>
	</tr>
	<tr>		
	    <td height="18px" align="center" class="MenuItem"><a href="m004q01.asp" target="ContactBodyFrame">Browse All</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" class="MenuItem"><a href="m004s0101.asp" target="ContactBodyFrame">Quick Search</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" class="MenuItem"><a href="m004s0102.asp" target="ContactBodyFrame">Advanced Search</a></td>
	</tr>
	<tr>		
	    <td height="18px" align="center" class="MenuItem"><a href="javascript: openWindow('m004a0101.asp');" target="ContactBodyFrame">New Contact</a></td>
	</tr>
</table>
</body>
</html>