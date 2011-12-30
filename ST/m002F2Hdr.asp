<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Staff Menu</title>
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
      <td nowrap" align="center"><a href="javascript: top.window.close();"><img src="../i/tn_staff_01.jpg" ALT="Return to Main Menu." width="81" height="60" border=0></a></td>
    </tr>
    <tr height="18px">
      <td nowrap align="center" class="MenuItem" width="120"><a href="m002d01.asp" target="StaffBodyFrame">Desktop</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="m002q01.asp" target="StaffBodyFrame">Browse All</a></td>
    </tr>
<!--
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem">Quick Search</td>
    </tr>
-->
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="javascript: openWindow('m002a0101.asp','wQA06');">New Staff</a></td>
    </tr>
  </table>
<!--</div>-->
</body>
</html>