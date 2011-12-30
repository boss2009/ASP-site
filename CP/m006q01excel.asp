<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsCompany__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsCompany__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsCompany__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsCompany__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsCompany__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsCompany__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,"+rsCompany__inspSrtBy.replace(/'/g, "''")+","+rsCompany__inspSrtOrd.replace(/'/g, "''")+",'"+rsCompany__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();
%>
<html>
<head>
	<title>Organization - Browse</title>
</head>
<body>
<table>
	<tr> 
		<th>Organization Name</th>
		<th>Type</th>
		<th>Address</th>
		<th>City</th>
		<th>Province/State</th>
		<th>Country</th>
		<th>Phone Number</th>
	</tr>
<% 
while (!rsCompany.EOF) { 
%>
	<tr> 
        <td><%=(rsCompany.Fields.Item("chvCompany_Name").Value)%></td>
        <td><%=(rsCompany.Fields.Item("chvWork_type_desc").Value)%></td>
        <td><%=(rsCompany.Fields.Item("chvAddress").Value)%></td>
        <td><%=(rsCompany.Fields.Item("chvCity").Value)%></td>
        <td><%=(rsCompany.Fields.Item("chrprvst_abbv").Value)%></td>
        <td><%=(rsCompany.Fields.Item("chvcntry_name").Value)%></td>
		<td><%=FormatPhoneNumber(rsCompany.Fields.Item("chvPhone_Type_1").Value,rsCompany.Fields.Item("chvPhone1_Arcd").Value,rsCompany.Fields.Item("chvPhone1_Num").Value,rsCompany.Fields.Item("chvPhone1_Ext").Value,rsCompany.Fields.Item("chvPhone_Type_2").Value,rsCompany.Fields.Item("chvPhone2_Arcd").Value,rsCompany.Fields.Item("chvPhone2_Num").Value,rsCompany.Fields.Item("chvPhone2_Ext").Value,rsCompany.Fields.Item("chvPhone_Type_3").Value,rsCompany.Fields.Item("chvPhone3_Arcd").Value,rsCompany.Fields.Item("chvPhone3_Num").Value,rsCompany.Fields.Item("chvPhone3_Ext").Value,rsCompany.Fields.Item("chvPhone3_Ext").Value)%></td>
	</tr>
<%
	rsCompany.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsCompany.Close();
%>