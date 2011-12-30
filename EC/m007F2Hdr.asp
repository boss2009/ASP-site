<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<html>
<head>
	<title>Equipment Class Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="javascript">
	function openWindow(page){
		if (page!='nothing') loadingstatus = window.open(page, "", "width=240,height=100,scrollbars=0,left=300,top=200,status=0");
	}

	function closeWindow(){
		if (loadingstatus!="undefined") loadingstatus.close();
	}
	</script>
</head>
<body>
<!--<div class="TestPanel" style="height: 450px">-->
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_inv_class_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m007d01.asp" target="EquipmentClassBodyFrame">Desktop</a></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m007s0101.asp" target="EquipmentClassBodyFrame">Quick Search</a></td>
	</tr>
<!--
	<tr nowrap>
		<td><a href="m007q0101.asp" target="EquipmentClassBodyFrame" onClick="openWindow('../loading.html');">Tree View</a></td>
	</tr>
-->
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m007q01lw.asp" target="EquipmentClassBodyFrame">Browse All</a></td>
	</tr>
<%
if (Session("MM_UserAuthorization") >= 5) {
%>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="m007t01.asp" target="EquipmentClassBodyFrame">Transfer Sub Abstract Class</a></td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="m007t02.asp" target="EquipmentClassBodyFrame">Transfer Concrete Class</a></td>
	</tr>		
<% 
} 
%>
</table>
<!--</div>-->
</body>
</html>