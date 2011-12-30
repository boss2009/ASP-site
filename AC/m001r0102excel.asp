<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
if(String(Request.QueryString("chvFilter")) != "") { 
  rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_adtclnt_srvnote_rpt_summary('"+ rsClient__chvFilter.replace(/'/g, "''") + "',0)}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
<html>
<head>
	<title>Funding Source, Service, Date Range Summary Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table>
	<tr> 
        <td>ASP ID</td>
        <td><%=rsClient.Fields.Item("chvSrvCd_1").Value%></td>
        <td><%=rsClient.Fields.Item("chvSrvCd_2").Value%></td>
        <td><%=rsClient.Fields.Item("chvSrvCd_3").Value%></td>
        <td><%=rsClient.Fields.Item("chvSrvCd_4").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_5").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_6").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_7").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_8").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_9").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_10").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_11").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_12").Value%></td>				
        <td><%=rsClient.Fields.Item("chvSrvCd_13").Value%></td>						
        <td>Total Service</td>						
    </tr>
<% 
rsClient.MoveNext();
var count_1 = 0
var count_2 = 0
var count_3 = 0
var count_4 = 0
var count_5 = 0
var count_6 = 0
var count_7 = 0
var count_8 = 0
var count_9 = 0
var count_10 = 0
var count_11 = 0
var count_12 = 0
var count_13 = 0
var total = 0
while (!rsClient.EOF) { 
%>
    <tr> 
        <td><%=(rsClient.Fields.Item("intAdult_id").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_1").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_2").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_3").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_4").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_5").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_6").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_7").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_8").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_9").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_10").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_11").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_12").Value)%>&nbsp;</td>
        <td><%=(rsClient.Fields.Item("intSrvCd_13").Value)%>&nbsp;</td>		
        <td><%=(rsClient.Fields.Item("intTtlSrvCd").Value)%>&nbsp;</td>
    </tr>
<%
	count_1 += rsClient.Fields.Item("intSrvCd_1").Value
	count_2 += rsClient.Fields.Item("intSrvCd_2").Value
	count_3 += rsClient.Fields.Item("intSrvCd_3").Value
	count_4 += rsClient.Fields.Item("intSrvCd_4").Value
	count_5 += rsClient.Fields.Item("intSrvCd_5").Value
	count_6 += rsClient.Fields.Item("intSrvCd_6").Value
	count_7 += rsClient.Fields.Item("intSrvCd_7").Value
	count_8 += rsClient.Fields.Item("intSrvCd_8").Value
	count_9 += rsClient.Fields.Item("intSrvCd_9").Value
	count_10 += rsClient.Fields.Item("intSrvCd_10").Value
	count_11 += rsClient.Fields.Item("intSrvCd_11").Value
	count_12 += rsClient.Fields.Item("intSrvCd_12").Value
	count_13 += rsClient.Fields.Item("intSrvCd_13").Value	
	total += rsClient.Fields.Item("intTtlSrvCd").Value
	rsClient.MoveNext();
}
%>
	<tr>
		<td>Total</td>
        <td><%=count_1%>&nbsp;</td>
        <td><%=count_2%>&nbsp;</td>
        <td><%=count_3%>&nbsp;</td>
        <td><%=count_4%>&nbsp;</td>
        <td><%=count_5%>&nbsp;</td>
        <td><%=count_6%>&nbsp;</td>
        <td><%=count_7%>&nbsp;</td>
        <td><%=count_8%>&nbsp;</td>
        <td><%=count_9%>&nbsp;</td>
        <td><%=count_10%>&nbsp;</td>
        <td><%=count_11%>&nbsp;</td>
        <td><%=count_12%>&nbsp;</td>
        <td><%=count_13%>&nbsp;</td>		
        <td><%=total%>&nbsp;</td>
	</tr>
</table>
</body>
</html>
<%
rsClient.Close();
%>