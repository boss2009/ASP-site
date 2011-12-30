<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsEquipmentClass__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsEquipmentClass__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsEquipmentClass__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsEquipmentClass__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}
var rsEquipmentClass__chvFilter = "b.chvSbjTotax <> '3'";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsEquipmentClass__chvFilter = String(Request.QueryString("chvFilter"));
}
var rsEquipmentClass = Server.CreateObject("ADODB.Recordset");
rsEquipmentClass.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentClass.Source = "{call dbo.cp_EC_Eqp_Class("+ rsEquipmentClass__inspSrtBy.replace(/'/g, "''") + ","+ rsEquipmentClass__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsEquipmentClass__chvFilter.replace(/'/g, "''") + "')}";
rsEquipmentClass.CursorType = 0;
rsEquipmentClass.CursorLocation = 2;
rsEquipmentClass.LockType = 3;
rsEquipmentClass.Open();
%>
<html>
<head>
	<title>Equipment Class - Browse</title>
</head>
<body>
<table>
    <tr> 
		<th class="headrow">Class Name</th>
		<th class="headrow">Class Type</th>
		<th class="headrow">Class ID</th>	  
    </tr>
<% 
while (!rsEquipmentClass.EOF) {
%>
    <tr>
<% 
	switch(rsEquipmentClass.Fields.Item("chrClass_Type").Value){ 
		case 'A': 
%>
		<td><%=(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%></td>	
		<td>Abstract</td>
<%
		break;
		case 'S':
%>
		<td><%=(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%></td>	
		<td>Sub Abstract</td>
<%
		break;
		case 'C':
%>
		<td><%=(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%></td>	
		<td>Concrete</td>
<%
		break;
	}
%>
		<td><%=ZeroPadFormat(rsEquipmentClass.Fields.Item("insEquip_Class_id").Value,8)%></td>		
    </tr>
<%
	rsEquipmentClass.MoveNext();
}
%>
</table>	
</body>
</html>
<%
rsEquipmentClass.Close();
%>
