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
%>
<html>
<head>
	<title>Shipping Information Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr>
    	<td><a href="m009e0401.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_Dtl_id=<%=rsEquipmentService.Fields.Item("intShip_Dtl_id").Value%>" target="SubBodyFrame">Shipping Method</a><%if (rsEquipmentService.Fields.Item("intShip_Dtl_id").Value!=null) Response.Write(" | ");%></td>
		<%
		if (rsEquipmentService.Fields.Item("intShip_Dtl_id").Value!=null) {
		%>		
		<td><a href="m009e0402.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_Dtl_id=<%=rsEquipmentService.Fields.Item("intShip_Dtl_id").Value%>" target="SubBodyFrame">Shipping Address</a> | </td>
		<td><a href="m009e0403.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_Dtl_id=<%=rsEquipmentService.Fields.Item("intShip_Dtl_id").Value%>" target="SubBodyFrame">Shipping Schedule</a></td>
		<%
		}
		%>
	</tr>
</table>
</body>
</html>