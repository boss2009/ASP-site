<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_Loan_Request("+ Request.QueryString("intAdult_id") + ")}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var rsLoanSummary = Server.CreateObject("ADODB.Recordset");
rsLoanSummary.ActiveConnection = MM_cnnASP02_STRING;
rsLoanSummary.Source = "{call dbo.cp_loan_hstry_summary("+ Request.QueryString("intAdult_id") + ",0)}";
rsLoanSummary.CursorType = 0;
rsLoanSummary.CursorLocation = 2;
rsLoanSummary.LockType = 3;
rsLoanSummary.Open();
%>
<html>
<head>
	<title>Loan Summary</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Loan Summary</h5>
<hr>
<% 
while (!rsLoanSummary.EOF) { 
%>
<b>Loan Request ID: <%=(rsLoanSummary.Fields.Item("intLoan_Req_id").Value)%></b>
<table cellpadding="1" cellspacing="2" class="MTable">
    <tr> 
		<th class="headrow" valign="top" align="left" width="300">Inventory Name</th>
		<th class="headrow" valign="top" align="left">Inventory ID</th>								
		<th class="headrow" valign="top" align="left">Date Processed</th>
		<th class="headrow" valign="top" align="left">Date Shipped</th>
		<th class="headrow" valign="top" align="left">Date Returned</th>						
    </tr>
<%
while (!rsLoan.EOF) {
	if (rsLoan.Fields.Item("intLoan_Req_id").Value==rsLoanSummary.Fields.Item("intLoan_Req_id").Value) {
%>
    <tr> 
		<td valign="top"><%=(rsLoan.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>
		<td nowrap valign="top" align="center"><%=ZeroPadFormat(rsLoan.Fields.Item("intEquip_Set_id").Value,8)%></td>						
		<td nowrap valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsProcess").Value)%></td>
		<td nowrap valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsDate_Shipped").Value)%></td>
		<td nowrap valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsReturn").Value)%></td>
    </tr>
<%
	}
	rsLoan.MoveNext();
}
%>		
</table><br>
<b>Total loan cost (excluding taxes and shipping): <%=FormatCurrency(rsLoanSummary.Fields.Item("fltPCost").Value)%><br>
Total cost of equipment still on loan: <%=FormatCurrency(rsLoanSummary.Fields.Item("fltPCost_Onloan").Value)%><br><br></b>
<%
	rsLoan.MoveFirst();
	rsLoanSummary.MoveNext();
}
%>
<hr>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsLoanSummary.Close();
%>