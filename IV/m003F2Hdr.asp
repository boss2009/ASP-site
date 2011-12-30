<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Inventory Menu</title>
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
	<tr>
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_inventory_01.jpg" ALT="Return to Main Menu." width="80" height="60" border=0></a></div></td>
	</tr>
	<tr> 
		<td height="18px" align="center" class="MenuItem" width="120"><a href="m003d01.asp" target="InventoryBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr> 		
	    <td height="18px" align="center" class="MenuItem"><a href="m003q02.asp" target="InventoryBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" class="MenuItem"><a href="m003s0101.asp" target="InventoryBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" class="MenuItem"><a href="m003s0102.asp" target="InventoryBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" class="MenuItem"><a href="m003r01menu.asp" target="InventoryBrowseRightFrame">Reports</a></td>
	</tr>	
	<tr>
		<td height="18px" align="center" class="MenuItem"><a href="javascript: openWindow('m003a0101.asp','w003A01');">New Inventory</a></td>
	</tr>	
	<tr>
		<td height="18px" align="center" class="MenuItem"></td>
	</tr>	
	<tr>
		<td height="18px" align="center" class="MenuItem"><a href="m003s0103.asp" target="InventoryBrowseRightFrame">Class Search</a></td>
	</tr>
</table>
</body>
</html>