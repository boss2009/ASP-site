<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Client Menu</title>
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
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/tn_client_01.jpg" ALT="Return to Main Menu." width="81" height="60"></a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem" width="120"><a href="m001d01.asp?intMOD_id=1" target="AdultClientBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001q01.asp" target="AdultClientBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001s0101.asp" target="AdultClientBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001s0102.asp" target="AdultClientBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001s0103.asp" target="AdultClientBrowseRightFrame">Power Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001s01menu.asp" target="AdultClientBrowseRightFrame">Preset Queries</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m001r01menu.asp" target="AdultClientBrowseRightFrame">Reports</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="javascript: openWindow('m001a0101a.asp','wQA01');">New Client</a></td>
	</tr>
</table>
</body>
</html>