<!--------------------------------------------------------------------------
* File Name: m009q02excel.asp
* Title: Equipment Service - Browse
* Main SP: cp_Get_eqp_srv
* Description: This page lists equipment services resulted from a search and
* export to excel.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsEquipmentService__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsEquipmentService__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsEquipmentService__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsEquipmentService__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsEquipmentService__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsEquipmentService__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_Get_Eqp_Srv2(0,"+rsEquipmentService__inspSrtBy+","+rsEquipmentService__inspSrtOrd+",'"+rsEquipmentService__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();
%>
<html>
<head>
	<title>Equipment Service - Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="2" cellspacing="1">
    <tr> 
        <th nowrap class="headrow" align="left">Equip. Serv. #.</th>
        <th nowrap class="headrow" align="left">Inventory ID</th>
        <th nowrap class="headrow" align="left">Inventory Name</th>
        <th nowrap class="headrow" align="left">Inventory Status</th>
        <th nowrap class="headrow" align="left">Repair Status</th>
        <th nowrap class="headrow" align="left">Date Requested</th>
        <th nowrap class="headrow" align="left">Date Completed</th>
        <th nowrap class="headrow" align="left">Repaired By</th>
        <th nowrap class="headrow" align="left">Service</th>
        <th nowrap class="headrow" align="left">Reason for Repair</th>
   </tr>
<% 
while (!rsEquipmentService.EOF) { 
%>
   <tr> 
        <td valign="top" align="left" nowrap><%=ZeroPadFormat(rsEquipmentService.Fields.Item("intEquip_Srv_id").Value, 8)%></td>
        <td valign="top" align="left" nowrap><%=ZeroPadFormat(rsEquipmentService.Fields.Item("intEquip_Set_id").Value, 8)%></td>
        <td valign="top" align="left" nowrap><%=Truncate(rsEquipmentService.Fields.Item("chvInventory_Name").Value,40)%>&nbsp;</td>
        <td valign="top" align="left" nowrap><%=(rsEquipmentService.Fields.Item("chvIvtry_Status").Value)%>&nbsp;</td>
        <td valign="top" align="left" nowrap><%=(rsEquipmentService.Fields.Item("chvRepair_Status").Value)%>&nbsp;</td>
        <td valign="top" align="center" nowrap><%=FilterDate(rsEquipmentService.Fields.Item("dtsRequested_date").Value)%>&nbsp;</td>
        <td valign="top" align="center" nowrap><%=FilterDate(rsEquipmentService.Fields.Item("dtsCompleted_Date").Value)%>&nbsp;</td>
        <td valign="top" align="left" nowrap><%=rsEquipmentService.Fields.Item("chvRepaired_by").Value%>&nbsp;</td>
        <td valign="top" align="center" nowrap><%=(rsEquipmentService.Fields.Item("insSrv_hrs").Value)%>Hr:<%=(rsEquipmentService.Fields.Item("insSrv_minutes").Value)%>Min&nbsp;</td>
        <td valign="top" align="left"><%=rsEquipmentService.Fields.Item("chrReason_Repair").Value%>&nbsp;</td>
	</tr>
<%
	rsEquipmentService.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsEquipmentService.Close();
%>