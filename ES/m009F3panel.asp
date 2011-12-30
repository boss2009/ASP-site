<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",0,0,'',1,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();

var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(10)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Equipment Service Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="JavaScript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}
</script>
</head>
<body onLoad="window.focus();">
<table align="center" cellspacing="0" width="130">
	<tr>
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/tn_service_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></td>
	</tr>
    <tr>
		<td height="18px" class="MenuItem" align="center"><a href="m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="BodyFrame">General Information</a></td>
    </tr>
    <tr>
		<td height="18px" class="MenuItem" align="center"><a href="m009e0201.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="BodyFrame">Service Requested</a></td>
    </tr>
    <tr>
		<td height="18px" class="MenuItem" align="center"><a href="m009FS03.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="BodyFrame">Service Performed</a></td>
    </tr>
    <tr>
		<td height="18px" class="MenuItem" align="center"><a href="m009FS04.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_dtl_id=<%=rsEquipmentService.Fields.Item("intShip_dtl_id").Value%>" target="BodyFrame">Shipping Information</a></td>
    </tr>
    <tr>
		<td height="18px" class="MenuItem" align="center"><a href="m009e0501.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="BodyFrame">EMail</a></td>
    </tr>
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m009a01j.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>','wj0901');" accesskey="D">Copy to DeskTop</a></td>
	</tr>
</table>
</body>
</html>
<%
rsFunction.Close();
%>