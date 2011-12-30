<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
Response.ContentType = "application/vnd.ms-excel"
%>
<%
var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "{call dbo.cp_Inventory_AdultClient_EduHstry("+Request.QueryString("insSchool_id")+",0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();
%>
<html>
<head>
	<title>Inventory Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th>Last Name</th>
		<th>First Name</th>
		<th>ASP ID</th>		
		<th>Inventory Name</th>		
		<th>Equipment Cost</th>		
		<th>Loan Status</th>		
		<th>Buyout Status</th>		
		<th>Date Returned</th>
    </tr>
<% 
while (!rsInventory.EOF) { 
%>
    <tr> 		
		<td><%=(rsInventory.Fields.Item("chvLst_name").Value)%></td>
		<td><%=(rsInventory.Fields.Item("chvFst_name").Value)%></td>		
		<td><%=(rsInventory.Fields.Item("intAdult_id").Value)%>&nbsp;</td>	  		
		<td><%=(rsInventory.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>		
		<td><%=(rsInventory.Fields.Item("fltEquip_Cost").Value)%>&nbsp;</td>		
		<td><%=(rsInventory.Fields.Item("chvLoan_Status").Value)%>&nbsp;</td>		
		<td><%=(rsInventory.Fields.Item("chvBuyoutStatus").Value)%>&nbsp;</td>		
		<td><%=(rsInventory.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
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
