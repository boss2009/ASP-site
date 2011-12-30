<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
if(String(Request.QueryString("chvFilter")) != "") { 
  rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_ac_srvnote_rpt_01A("+ Request.QueryString("intAdult_id") + ",'"+rsClient__chvFilter.replace(/'/g, "''")+"',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();
%>
<html>
<head>
	<title>Service Funding Source</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Service Funding Source</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="150">Service Date</th>	
		<th class="headrow" nowrap align="left">Service Code</th>
		<th class="headrow" nowrap align="left">Funded By</th>
    </tr>
<% 
while (!rsFundingSource.EOF) { 
%>
    <tr> 
		<td nowrap><%=FilterDate(rsFundingSource.Fields.Item("dtsRequest_Date").Value)%>&nbsp;</td>
		<td nowrap><%=(rsFundingSource.Fields.Item("chvSrv_Code").Value)%>&nbsp;</td>
		<td nowrap><%=(rsFundingSource.Fields.Item("chvFunding_Src").Value)%>&nbsp;</td>
    </tr>
<%
	rsFundingSource.MoveNext();
}
%>
</table>
<br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsFundingSource.Close();
%>