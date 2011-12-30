<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Report Menu</title>
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
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_report.jpg" ALT="Return to Main Menu." width="80" height="63"></a></div></td>
	</tr>
	<tr> 
		<td height="17" align="center" nowrap class="MenuItem" width="120"><a href="..\AC\m001r01menu.asp" target="RPBrowseRightFrame">client</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\SH\m012r0101.asp" target="RPBrowseRightFrame">Institutions</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\CP\m006r0101.asp" target="RPBrowseRightFrame">Organizations</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\IV\m003r01menu.asp" target="RPBrowseRightFrame">Inventory</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\PR\m014r0101.asp" target="RPBrowseRightFrame">Purchase Requisition</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\BD\m005r0101.asp" target="RPBrowseRightFrame">Equipment Bundles</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\LN\m008r0101.asp" target="RPBrowseRightFrame">Loan</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\BO\m010r0101.asp" target="RPBrowseRightFrame">BO Rpt 1</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="..\BO\m010r0201.asp" target="RPBrowseRightFrame">BO Rpt 2</a></td>
	</tr>
</table>
</body>
</html>