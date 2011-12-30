<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Loan Menu</title>
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
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_loan_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem" width="120"><a href="m008d01.asp" target="LoanBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008q01.asp" target="LoanBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008s0101.asp" target="LoanBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008s0201.asp" target="LoanBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008s0301.asp" target="LoanBrowseRightFrame">Power Search</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008r0101.asp" target="LoanBrowseRightFrame">Reports</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" class="MenuItem"><a href="m008s0401.asp" target="LoanBrowseRightFrame">Loan Equip<br>Request To Do</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m008c0101.asp" target="LoanBrowseRightFrame">Work Priority</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem"><a href="javascript: openWindow('m008a0101.asp','wQA01');">New Loan Request</a></td>
	</tr>
</table>
</body>
</html>