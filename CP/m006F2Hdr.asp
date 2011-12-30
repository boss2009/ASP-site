<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Organization Menu</title>
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
<!--<div class="TestPanel" style="height:476px;"> -->
<table align="center" cellspacing="0">
    <tr> 
		<td nowrap" align="center"><a href="javascript: top.window.close();"><img src="../i/tn_organization_02.jpg" ALT="Return to Main Menu." width="81" height="60" border=0></a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem" width="120"><a href="m006d01.asp" target="CompaniesBodyFrame">Desktop</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m006q01.asp" target="CompaniesBodyFrame">Browse All</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m006s0101.asp" target="CompaniesBodyFrame">Quick Search</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m006s0102.asp" target="CompaniesBodyFrame">Advanced Search</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m006r0101.asp" target="CompaniesBodyFrame">Reports</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="javascript: openWindow('m006a0101.asp','wQA06');">New Organization</a></td>
    </tr>
</table>
<!--</div>-->
</body>
</html>