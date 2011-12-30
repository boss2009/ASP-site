<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
Response.ContentType = "application/vnd.ms-excel"

var rsPILATStudent__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsPILATStudent__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsPILATStudent__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsPILATStudent__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsPILATStudent__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsPILATStudent__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStudent.Source = "{call dbo.cp_pilat_student(0,'','','','','',0,0,0,0,'',0,0,0,"+rsPILATStudent__inspSrtBy+","+rsPILATStudent__inspSrtOrd+",'"+rsPILATStudent__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsPILATStudent.CursorType = 0;
rsPILATStudent.CursorLocation = 2;
rsPILATStudent.LockType = 3;
rsPILATStudent.Open();
%>
<html>
<head>
	<title>Temp Student - Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table>
	<tr> 
    	<th>Last Name</th>
    	<th>First Name</th>		
        <th>SIN</th>
        <th>Primary Disability</th>
        <th>Status</th>
	</tr>
<% 
while (!rsPILATStudent.EOF) { 
%>
	<tr> 
        <td><%=(rsPILATStudent.Fields.Item("chvLst_Name").Value)%></td>
		<td><%=(rsPILATStudent.Fields.Item("chvFst_Name").Value)%></td>
        <td><%=(rsPILATStudent.Fields.Item("chrSIN_no").Value)%></td>
        <td><%=(rsPILATStudent.Fields.Item("chvDisability").Value)%></td>
        <td><%=(rsPILATStudent.Fields.Item("chvStdnt_Status").Value)%></td>
	</tr>
<%
	rsPILATStudent.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsPILATStudent.Close();
%>