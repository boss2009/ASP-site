<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Equipment Bundles Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=480,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>	
</head>
<body>
<table align="center" cellspacing="0">
    <tr> 
      <td nowrap" align="center"><a href="javascript: top.window.close();"><img src="../i/tn_equip_bundle_01.jpg" ALT="Return to Main Menu." width="80" height="60" border=0></a></td>
    </tr>
    <tr height="18px">
      <td nowrap align="center" class="MenuItem" width="120"><a href="m005d01.asp" target="EquipmentBundleBodyFrame">Desktop</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="m005q01.asp" target="EquipmentBundleBodyFrame">Browse All</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="m005s0101.asp" target="EquipmentBundleBodyFrame">Quick Search</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="m005s0102.asp" target="EquipmentBundleBodyFrame">Advanced Search</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="m005r0101.asp" target="EquipmentBundleBodyFrame">Reports</a></td>
    </tr>
    <tr height="18px"> 
      <td nowrap align="center" class="MenuItem"><a href="javascript: openWindow('m005a0101.asp','wQA06');">New Equipment Bundle</a></td>
    </tr>
</table>
</body>
</html>