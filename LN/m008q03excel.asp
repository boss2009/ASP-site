<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsLoan__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
	rsLoan__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsLoan__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
	rsLoan__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsLoan__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
	rsLoan__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request4(0,0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,"+rsLoan__inspSrtBy+","+rsLoan__inspSrtOrd+",'"+rsLoan__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();
%>
<html>
<head>
	<title>Loan - Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h3>Loan - Browse</h3>
<table>
	<tr> 
		<th>Loan Description</th>
		<th>Disability</th>
		<th>Case Manager</th>
		<th>Loan Type</th>
		<th>Loan Status</th>
		<th>Date Processed</th>
		<th>Delivery Date</th>
    </tr>
<% 
while (!rsLoan.EOF) { 
%>
    <tr> 
		<td><%=(rsLoan.Fields.Item("chvLoan_name").Value)%>&nbsp;</td>
		<td><%=(rsLoan.Fields.Item("chvDisability").Value)%>&nbsp;</td>
		<td><%=(rsLoan.Fields.Item("chvCase_Manager").Value)%>&nbsp;</td>
		<td><%=(rsLoan.Fields.Item("chvLoan_Type_id").Value)%>&nbsp;</td>
		<td><%=(rsLoan.Fields.Item("chvLoan_Status_id").Value)%>&nbsp;</td>
		<td><%=FilterDate(rsLoan.Fields.Item("dtsProcessed_date").Value)%>&nbsp;</td>
		<td><%=FilterDate(rsLoan.Fields.Item("dtsDelivery_date").Value)%>&nbsp;</td>
    </tr>
<%
	rsLoan.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsLoan.Close();
%>