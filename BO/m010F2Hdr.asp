<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Buyout Menu</title>
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
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_buyout_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem" width="120"><a href="m010d01.asp" target="BuyoutBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010q01.asp" target="BuyoutBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010s0101.asp" target="BuyoutBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010s0201.asp" target="BuyoutBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010r0101.asp" target="BuyoutBrowseRightFrame">Report 1</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010r0201.asp" target="BuyoutBrowseRightFrame">Report 2</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010s0301.asp" target="BuyoutBrowseRightFrame">Buyout Equip<br>Request To Do</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m010c0101.asp" target="BuyoutBrowseRightFrame">Work Priority</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="javascript: openWindow('m010a0101.asp','wQA10');">New Buyout Request</a></td>
	</tr>
</table>
</body>
</html>