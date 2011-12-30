<!--------------------------------------------------------------------------
* File Name: m010q05excel.asp
* Title: Buyout - Browse
* Main SP: cp_report_buyout_request2
* Description: This page lists buyouts resulted from a search and exports
* to excel.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%Response.ContentType = "application/vnd.ms-excel"%>
<%
var rsBuyout__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsBuyout__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsBuyout__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsBuyout__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsBuyout__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsBuyout__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_report_buyout_request2("+rsBuyout__inspSrtBy+","+rsBuyout__inspSrtOrd+",'"+rsBuyout__chvFilter.replace(/'/g, "''")+"',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();
%>

<html>
<head>
	<title>Buyout - Browse</title>
</head>
<body>
<table>
	<tr> 
		<th>Buyout ID</th>
		<th>Funding Source</th>
		<th>Buyer Name</th>
		<th>Case Manager</th>
		<th>Is Backordered</th>
		<th>Buyout Status</th>
		<th>Buyout Process</th>
	</tr>
<% 
while (!rsBuyout.EOF) { 
%>
	<tr> 
        <td nowrap><%=rsBuyout.Fields.Item("intBuyout_Req_id").Value%></td>
        <td nowrap><%=rsBuyout.Fields.Item("chvfunding_source_name").Value%>&nbsp;</td>		
        <td nowrap><%=((rsBuyout.Fields.Item("insEq_user_type").Value==3)?rsBuyout.Fields.Item("chvLst_Name").Value+", "+rsBuyout.Fields.Item("chvFst_Name").Value:rsBuyout.Fields.Item("chvSchool_Name").Value)%>&nbsp;</td>		
        <td nowrap><%=FilterDate(rsBuyout.Fields.Item("chvCaseManager").Value)%>&nbsp;</td>
        <td nowrap><%=rsBuyout.Fields.Item("bitIsBack_Ordered").Value%>&nbsp;</td>
        <td nowrap><%=rsBuyout.Fields.Item("chvBO_Status").Value%>&nbsp;</td>
        <td nowrap><%=rsBuyout.Fields.Item("chvBO_Process").Value%>&nbsp;</td>
	</tr>
<%
	rsBuyout.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsBuyout.Close();
%>