<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_Request("+ Request.QueryString("intAdult_id") + ")}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var rsBuyoutSummary = Server.CreateObject("ADODB.Recordset");
rsBuyoutSummary.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutSummary.Source = "{call dbo.cp_buyout_hstry_summary("+ Request.QueryString("intAdult_id") + ",1,1,1,0)}";
rsBuyoutSummary.CursorType = 0;
rsBuyoutSummary.CursorLocation = 2;
rsBuyoutSummary.LockType = 3;
rsBuyoutSummary.Open();
%>
<html>
<head>
	<title>Buyout Summary</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Buyout Summary</h5>
<hr>
<% 
while (!rsBuyoutSummary.EOF) { 
%>
<b>Buyout ID: <%=(rsBuyoutSummary.Fields.Item("intBuyout_req_id").Value)%></b>
<table cellspacing="1" cellpadding="2" class="Mtable">
    <tr> 
		<th class="headrow" valign="top" align="left" width="300">Inventory Name</th>
		<th class="headrow" valign="top" align="left">Inventory ID</th>								
		<th class="headrow" valign="top" align="left">Date Processed</th>
		<th class="headrow" valign="top" align="left">Date Returned</th>						
    </tr>
<%
while (!rsBuyout.EOF) {
	if (rsBuyout.Fields.Item("intBuyout_req_id").Value==rsBuyoutSummary.Fields.Item("intBuyout_req_id").Value) {
%>
    <tr> 
		<td valign="top"><%=(rsBuyout.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>
		<td nowrap valign="top" align="center"><%=ZeroPadFormat(rsBuyout.Fields.Item("intEquip_Set_id").Value,8)%></td>						
		<td nowrap valign="top" align="center"><%=FilterDate(rsBuyout.Fields.Item("dtsProcess").Value)%></td>
		<td nowrap valign="top" align="center"><%=FilterDate(rsBuyout.Fields.Item("dtsReturn").Value)%></td>	
    </tr>
<%
	}
	rsBuyout.MoveNext();
}
%>		
</table><br>
<b>
Total sold price (excluding taxes and shipping): <%=FormatCurrency(rsBuyoutSummary.Fields.Item("fltEqp_Sold_price").Value)%><br>
Total sold price (including taxes and shipping): <%=FormatCurrency(rsBuyoutSummary.Fields.Item("fltEqp_Sold_price_TaxShip").Value)%><br><br>
</b>
<%
	rsBuyout.MoveFirst();
	rsBuyoutSummary.MoveNext();
}
%>
<hr>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsBuyoutSummary.Close();
%>