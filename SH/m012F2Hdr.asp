<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Institution Menu</title>
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
<!--<div class="TestPanel" style="height:476px;">-->
<table align="center" cellspacing="0">
    <tr>
		<td nowrap" align="center"><a href="javascript: top.window.close();"><img src="../i/tn_institution_01.jpg" ALT="Return to Main Menu." width="80" height="70" border=0></a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem" width="120"><a href="m012d01.asp" target="InstitutionBodyFrame">Desktop</a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem"><a href="m012q01.asp" target="InstitutionBodyFrame">Browse All</a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem"><a href="m012s0101.asp" target="InstitutionBodyFrame">Quick Search</a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem"><a href="m012s0102.asp" target="InstitutionBodyFrame">Advanced Search</a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem"><a href="m012r0101.asp" target="InstitutionBodyFrame">Report</a></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem"><a href="javascript: openWindow('m012a0101.asp','wQA12');">New Institution</a></td>
    </tr>
</table>
<!--</div>-->
</body>
</html>