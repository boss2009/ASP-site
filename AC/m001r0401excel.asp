<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" --> 
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsClient__inspSrtBy = "1";
if(String(Request.QueryString("SortByColumn")) != "undefined") { 
  rsClient__inspSrtBy = String(Request.QueryString("SortByColumn"));
}

var rsClient__inspSrtOrd = "0";
if(String(Request.QueryString("OrderBy")) != "undefined") { 
  rsClient__inspSrtOrd = String(Request.QueryString("OrderBy"));
}

var rsClient__chvFilter = "chvContactName  = ''";
if(String(Request.QueryString("MM_param")) != "undefined") { 
	rsClient__chvFilter = Request.QueryString("MM_param");
} else{
	rsClient__chvFilter = " insRtnship_id= " + String(Request.QueryString("insRtnship_id"));
}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_AC_Contact("+ rsClient__inspSrtBy.replace(/'/g, "''") + ","+ rsClient__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
<html>
<head>
	<title><%=((String(Request.QueryString("flag")) == "1")?"Contact Excel Report":"Contact Relationship Excel Report")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table>
    <tr> 
		<td>Last Name</td>
		<td>First Name</td>
		<td>Referral Date</td>
		<td>Re-Referral Date</td>
		<td>Disability</td>
		<td>Region</td>
		<td>Status</td>
		<td>SIN</td>
		<td>Gender</td>
		<td>Age</td>
    </tr>
<% 
while (!rsClient.EOF) { 
%>
    <tr> 
		<td><%=(rsClient.Fields.Item("chvLst_Name").Value)%></td>
		<td><%=(rsClient.Fields.Item("chvFst_Name").Value)%></td>
		<td><%=FilterDate(rsClient.Fields.Item("dtsRefral_date").Value)%></td>
		<td><%=FilterDate(rsClient.Fields.Item("dtsRe_refral_date").Value)%></td>
		<td><%=(rsClient.Fields.Item("chvDisability").Value)%></td>
		<td><%=(rsClient.Fields.Item("chvRegion").Value)%></td>
		<td><%=(rsClient.Fields.Item("chvStatus").Value)%></td>
		<td><%=(rsClient.Fields.Item("chrSIN_no").Value)%></td>
		<td><%=(rsClient.Fields.Item("chrGender").Value)%></td>
		<td><%=(rsClient.Fields.Item("intAge").Value)%></td>
    </tr>
<%
	rsClient.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsClient.Close();
%>