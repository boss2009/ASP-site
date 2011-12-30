<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var CheckEquipmentService = Server.CreateObject("ADODB.Command");

// + Nov.11.2005
var intTmp = Request.QueryString("intEquip_srv_id");
if (intTmp == null) intTmp = 0;

CheckEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
CheckEquipmentService.CommandText = "dbo.cp_chk_eqpsrv_type_A";
CheckEquipmentService.CommandType = 4;
CheckEquipmentService.CommandTimeout = 0;
CheckEquipmentService.Prepared = true;
CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("RETURN_VALUE", 3, 4));
CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@intEquip_srv_id", 3, 1,10000,Request.QueryString("intEquip_srv_id")));
CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@insRtnFlag", 2, 2));
CheckEquipmentService.Execute();
%>
<html>
<head>
	<title>General Information Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m009e0101.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">Inventory Information</a> | </td>
<%
switch (CheckEquipmentService.Parameters.Item("@insRtnFlag").Value) {
	case 8:
%>
		<td><a href="m009e0102.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">Individual User</a> | </td>
<%
	break;
	case 9:
%>
		<td><a href="m009e0103.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">Institution User</a> | </td>		
<%
	break;
}
%>
		<td><a href="m009e0104.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">Repair Status</a></td>
	</tr>
</table>
</body>
</html>