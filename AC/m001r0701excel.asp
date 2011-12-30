<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsClient__chvFilter = "dtsRefral_date >= '01/01/1900'";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Adult_Client3D("+ Request.QueryString("inspSrtBy") + ","+ Request.QueryString("inspSrtOrd") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
<html>
<head>
	<title>Client - Browse</title>
</head>
<body>
<table>
    <tr> 
		<th>Last Name</th>
		<th>First Name</th>
		<th>ASP ID</th>
		<th>SIN</th>
		<th>Status</th>
		<th>Phone Number</th>		
		<th>Case Manager</th>
		<th>Primary Disability</th>		
		<th>Referral Date</th>
		<th>Re-referral Date</th>
    </tr>
<% 
while (!rsClient.EOF) { 
%>
    <tr> 
		<td><%=(rsClient.Fields.Item("chvLst_Name").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("chvFst_Name").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("intAdult_Id").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("chrSIN_no").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("chvStatus").Value)%>&nbsp;</td>		
		<td><%=(rsClient.Fields.Item("chvPhone_no").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("chvCaseManager").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("chvDisability").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
		<td><%=(rsClient.Fields.Item("dtsRe_refral_date").Value)%>&nbsp;</td>
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
