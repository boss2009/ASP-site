<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
Response.ContentType = "application/vnd.ms-excel"
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
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_02("+ rsInventory__inspSrtBy.replace(/'/g, "''") + ","+ rsInventory__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsInventory__chvFilter.replace(/'/g, "''") + "',0,0,0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();
%>
<html>
<head>
	<title>Inventory General Query</title>
</head>
<body>
<table>
    <tr> 
		<th nowrap class="headrow" align="left">Inventory Name</th>
		<th nowrap class="headrow" align="left">Inventory ID</th>
		<th nowrap class="headrow" align="left">Model Number</th>
		<th nowrap class="headrow" align="left">Serial Number</th>
		<th nowrap class="headrow" align="left">Requisition Number</th>
		<th nowrap class="headrow" align="left">Vendor</th>
		<th nowrap class="headrow" align="left">Current Status</th>
		<th nowrap class="headrow" align="left">Current User</th>
		<th nowrap class="headrow" align="left">Current Institution User</th>
		<th nowrap class="headrow" align="left">Inventory Cost</th>
		<th nowrap class="headrow" align="left">Sold Cost</th>
    </tr>    
<% 
while (!rsInventory.EOF) { 
%>
    <tr> 
		<td nowrap><%=(rsInventory.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>
		<td nowrap><%=ZeroPadFormat(rsInventory.Fields.Item("intBar_Code_no").Value,8)%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("chvModel_Number").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("chvSerial_Number").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("intRequisition_no").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("chvVendor_Name").Value%>&nbsp;</td>
		<td nowrap><%=(rsInventory.Fields.Item("chvEqp_Status").Value)%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("chvIdvUsr_Nm").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("chvInstitUsr_Nm").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("fltList_Unit_Cost").Value%>&nbsp;</td>
		<td nowrap><%=rsInventory.Fields.Item("fltPurchase_Cost").Value%>&nbsp;</td>
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