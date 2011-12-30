<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
Response.ContentType = "application/vnd.ms-excel" 

var rsBundle__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsBundle__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsBundle__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsBundle__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsBundle__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsBundle__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsBundle = Server.CreateObject("ADODB.Recordset");
rsBundle.ActiveConnection = MM_cnnASP02_STRING;
rsBundle.Source = "{call dbo.cp_Bundle2(0,'',0.0,0,1,1,'',0,"+Session("insStaff_id")+","+rsBundle__inspSrtBy.replace(/'/g, "''")+","+rsBundle__inspSrtOrd.replace(/'/g, "''")+",'"+rsBundle__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsBundle.CursorType = 0;
rsBundle.CursorLocation = 2;
rsBundle.LockType = 3;
rsBundle.Open();
%>
<html>
<head>
	<title>Equipment Bundle</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table>
	<tr> 
		<th>Equipment Bundle Name</th>
		<th>List Unit Cost</th>
		<th>Status</th>
	</tr>
<% 
while (!rsBundle.EOF) { 
%>
	<tr> 
		<td><%=(rsBundle.Fields.Item("chvName").Value)%></td>
		<td align="right"><%=FormatCurrency(rsBundle.Fields.Item("FltList_Unit_Cost").Value)%></td>
		<td><%=((rsBundle.Fields.Item("bitBundle_Status").Value=="1")?"Active":"Inactive")%></td>
	</tr>
<%
	rsBundle.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsBundle.Close();
%>