<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Temp Student Menu</title>
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
		<td nowrap" align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_student_01.jpg" ALT="Return to Main Menu." width="80" height="60" border="0"></a></div></td>
    </tr>
    <tr height="18px">
		<td nowrap align="center" class="MenuItem" width="120"><a href="m022d01.asp" target="PILATStudentBodyFrame">Desktop</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m022q01.asp" target="PILATStudentBodyFrame">Browse All</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="m022s0101.asp" target="PILATStudentBodyFrame">Quick Search</a></td>
    </tr>
    <tr height="18px"> 
		<td nowrap align="center" class="MenuItem"><a href="javascript: openWindow('m022a0101.asp','wQA22');">New Temp Student</a></td>
    </tr>
</table>
<!--</div>-->
</body>
</html>