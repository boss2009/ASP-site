<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Equipment Service Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<table align="center" cellspacing="0">
	<tr height="100">
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_service_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem" width="120"><a href="m009d01.asp" target="EquipmentServiceBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m009q01.asp" target="EquipmentServiceBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m009s0101.asp" target="EquipmentServiceBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m009s0102.asp" target="EquipmentServiceBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="javascript: openWindow('m009a0101.asp','wQA09');">New Equipment Service</a></td>
	</tr>
</table>
</body>
</html>