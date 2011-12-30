<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsInventory__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsInventory__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsInventory__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsInventory__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsInventory__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsInventory__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_05("+ rsInventory__inspSrtBy.replace(/'/g, "''") + ","+ rsInventory__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsInventory__chvFilter.replace(/'/g, "''") + "',0,0,0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();
%>
<html>
<head>
	<title>Inventory Loan Request Report</title>
	<meta http-equiv="Content-Type" content="application/vnd.ms-excel">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table>
    <tr> 
		<th>Inventory Name</th>
		<th>Inventory ID</th>
		<th>Model Number</th>
		<th>Serial Number</th>
		<th>Requisition Number</th>
		<th>Vendor</th>
		<th>Current Status</th>
		<th>Current User</th>
		<th>Current Institution</th>
		<th>Inventory Cost</th>
		<th>Sold Cost</th>
		<th>Loaned To</th>
		<th>Delivery Date</th>	  
    </tr>
<% 
while (!rsInventory.EOF) { 
%>
    <tr> 
		<td><%=rsInventory.Fields.Item("chvInventory_Name").Value%></td>
		<td><%=rsInventory.Fields.Item("intEquip_Set_id").Value%></td>
		<td><%=rsInventory.Fields.Item("chvModel_Number").Value%></td>
		<td><%=rsInventory.Fields.Item("chvSerial_Number").Value%></td>
		<td><%=rsInventory.Fields.Item("intRequisition_no").Value%></td>
		<td><%=rsInventory.Fields.Item("chvVendor_Name").Value%></td>
		<td><%=(rsInventory.Fields.Item("chvEqp_Status").Value)%></td>
		<td><%=rsInventory.Fields.Item("chvIdvUsr_Nm").Value%></td>
		<td><%=rsInventory.Fields.Item("chvInstitUsr_Nm").Value%></td>
		<td><%=rsInventory.Fields.Item("fltList_Unit_Cost").Value)%></td>
		<td><%=rsInventory.Fields.Item("fltPurchase_Cost").Value)%></td>      
		<td><%=(rsInventory.Fields.Item("chvLoaned_to").Value)%></td>      
		<td><%=(rsInventory.Fields.Item("dtsDlvy_date").Value)%></td>	  
    </tr>
<%
	rsInventory.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsInventory.Close();
%>
