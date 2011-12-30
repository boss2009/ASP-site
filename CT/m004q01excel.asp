<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsContact__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsContact__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsContact__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsContact__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsContact__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsContact__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_Contacts(0,0,'','','',0,0,0,"+rsContact__inspSrtBy.replace(/'/g, "''")+","+rsContact__inspSrtOrd.replace(/'/g, "''")+",'"+rsContact__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();
%>
<html>
<head>
	<title>Contact Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table>
    <tr> 
		<th nowrap>Last Name</th>
		<th nowrap>First Name</th>
		<th nowrap>Title</th>	  
		<th nowrap>Job Title</th>	  
		<th nowrap>Work Type</th>
    </tr>
<% 
while (!rsContact.EOF)) { 
%>
    <tr> 
      <td nowrap><%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
	  <td nowrap><%=(rsContact.Fields.Item("chvFst_Name").Value)%></td>
      <td nowrap><%=(rsContact.Fields.Item("chvtitle").Value)%></td>	  
      <td nowrap><%=(rsContact.Fields.Item("chvJob_Title").Value)%></td>	  
      <td nowrap><%=(rsContact.Fields.Item("chvWork_type_desc").Value)%></td>
    </tr>
<%
	rsContact.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsContact.Close();
%>
