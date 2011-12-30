<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
Response.ContentType = "application/vnd.ms-excel"
var rsInstitution__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsInstitution__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsInstitution__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsInstitution__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsInstitution__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsInstitution__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_query_school(0,"+rsInstitution__inspSrtBy+","+rsInstitution__inspSrtOrd+",'"+rsInstitution__chvFilter.replace(/'/g, "''")+"',0,0)}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();
%>
<html>
<head>
	<title>Institution Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table>
	<tr> 
    	<th>Institution Name</th>
        <th>Region</th>
        <th>School Type</th>
		<th>Number of Referrals</th>
		<th>Referral Date</th>
		<th>PILAT Status</th>
        <th>Is Main Campus</th>
	</tr>
<% 
while (!rsInstitution.EOF) { 
%>
	<tr> 
        <td><%=(rsInstitution.Fields.Item("chvSchool_Name").Value)%></td>
        <td><%=(rsInstitution.Fields.Item("chvRegion").Value)%></td>
        <td><%=(rsInstitution.Fields.Item("chvSchool_Type").Value)%></td>
        <td><%=(rsInstitution.Fields.Item("intRCnt").Value)%></td>		
        <td><%=FilterDate(rsInstitution.Fields.Item("dtsReferral_date").Value)%></td>		
        <td><%=(rsInstitution.Fields.Item("chvPILAT_Status").Value)%></td>			
        <td><%=(rsInstitution.Fields.Item("bitIs_MainCampus").Value)%></td>
	</tr>
<%
	rsInstitution.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsInstitution.Close();
%>