<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
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
var rsInventory_numRows = 0;
%>
<%
var Repeat1__numRows = 20;
var Repeat1__index = 0;
rsInventory_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventory_total = rsInventory.RecordCount;

// set the number of rows displayed on this page
if (rsInventory_numRows < 0) {            // if repeat region set to all records
  rsInventory_numRows = rsInventory_total;
} else if (rsInventory_numRows == 0) {    // if no repeat regions
  rsInventory_numRows = 1;
}

// set the first and last displayed record
var rsInventory_first = 1;
var rsInventory_last  = rsInventory_first + rsInventory_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventory_total != -1) {
  rsInventory_numRows = Math.min(rsInventory_numRows, rsInventory_total);
  rsInventory_first   = Math.min(rsInventory_first, rsInventory_total);
  rsInventory_last    = Math.min(rsInventory_last, rsInventory_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventory_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventory_total=0; !rsInventory.EOF; rsInventory.MoveNext()) {
    rsInventory_total++;
  }

  // reset the cursor to the beginning
  if (rsInventory.CursorType > 0) {
    if (!rsInventory.BOF) rsInventory.MoveFirst();
  } else {
    rsInventory.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventory_numRows < 0 || rsInventory_numRows > rsInventory_total) {
    rsInventory_numRows = rsInventory_total;
  }

  // set the first and last displayed record
  rsInventory_last  = Math.min(rsInventory_first + rsInventory_numRows - 1, rsInventory_total);
  rsInventory_first = Math.min(rsInventory_first, rsInventory_total);
}
%>
<html>
<head>
	<title>Inventory Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h3>Inventory Browse</h3>
<table>
    <tr> 
      <td colspan="4" align="left">Displaying <b><%=(rsInventory_total)%></b> Records.</td>
    </tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th nowrap class="headrow" align="left" width="280">Inventory Name</th>
		<th nowrap class="headrow" align="left">Inventory ID</th>
		<th nowrap class="headrow" align="left">Model Number</th>
		<th nowrap class="headrow" align="left">Serial Number</th>
		<th nowrap class="headrow" align="left">PR Number</th>
		<th nowrap class="headrow" align="left">Vendor</th>
		<th nowrap class="headrow" align="left">Current Status</th>
		<th nowrap class="headrow" align="left">Current User</th>
		<th nowrap class="headrow" align="left">Inventory Cost</th>
		<th nowrap class="headrow" align="left">Sold Cost</th>
		<th nowrap class="headrow" align="left">Loaned To</th>
		<th nowrap class="headrow" align="left">Delivery Date</th>
    </tr>
<% 
var total_sold = 0;
while (!rsInventory.EOF) { 
%>
    <tr> 
		<td valign="top" align="left"><%=Truncate(rsInventory.Fields.Item("chvInventory_Name").Value,40)%></td>
		<td valign="top" align="center" nowrap><%=ZeroPadFormat(rsInventory.Fields.Item("intEquip_Set_id").Value,8)%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("chvModel_Number").Value%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("chvSerial_Number").Value%></td>
		<td valign="top" align="center" nowrap><%=ZeroPadFormat(rsInventory.Fields.Item("intRequisition_no").Value,8)%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("chvVendor_Name").Value%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("chvEqp_Status").Value%></td>
		<td valign="top" align="left" nowrap><%=((rsInventory.Fields.Item("chvInstitUsr_Nm").Value=="")?rsInventory.Fields.Item("chvIdvUsr_Nm").Value:rsInventory.Fields.Item("chvInstitUsr_Nm").Value)%></td>
		<td valign="top" align="left" nowrap><%=FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value)%></td>
		<td valign="top" align="left" nowrap><%=FormatCurrency(rsInventory.Fields.Item("fltPurchase_Cost").Value)%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("chvLoaned_to").Value%></td>
		<td valign="top" align="left" nowrap><%=rsInventory.Fields.Item("dtsDlvy_date").Value%></td>
    </tr>
<%
	total_sold = total_sold + rsInventory.Fields.Item("fltList_Unit_Cost").Value;
	Repeat1__index++;
	rsInventory.MoveNext();
}
%>
	<tr>
		<td colspan="9"></td>
		<td align="right"><%=FormatCurrency(total_sold)%></td>
		<td cospan="2"></td>
	</tr>
</table>
</body>
</html>
<%
rsInventory.Close();
%>
