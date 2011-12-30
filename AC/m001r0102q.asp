<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
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
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
</head>
<body>
<h5>Funding Source, Service, Date Range Summary Report</h5>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
        <th nowrap class="headrow" align="left">ASP ID</th>
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_1").Value%></th>
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_2").Value%></th>
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_3").Value%></th>
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_4").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_5").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_6").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_7").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_8").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_9").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_10").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_11").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_12").Value%></th>				
        <th nowrap class="headrow" align="left"><%=rsClient.Fields.Item("chvSrvCd_13").Value%></th>						
        <th nowrap class="headrow" align="left">Total Service</th>						
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
        <td align="left"><%=(rsClient.Fields.Item("intAdult_id").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_1").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_2").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_3").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_4").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_5").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_6").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_7").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_8").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_9").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_10").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_11").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_12").Value)%>&nbsp;</td>
        <td align="center"><%=(rsClient.Fields.Item("intSrvCd_13").Value)%>&nbsp;</td>		
        <td align="center"><%=(rsClient.Fields.Item("intTtlSrvCd").Value)%>&nbsp;</td>
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
		<td align="left"><b>Total</b></td>
        <td align="center"><b><%=count_1%>&nbsp;</b></td>
        <td align="center"><b><%=count_2%>&nbsp;</b></td>
        <td align="center"><b><%=count_3%>&nbsp;</b></td>
        <td align="center"><b><%=count_4%>&nbsp;</b></td>
        <td align="center"><b><%=count_5%>&nbsp;</b></td>
        <td align="center"><b><%=count_6%>&nbsp;</b></td>
        <td align="center"><b><%=count_7%>&nbsp;</b></td>
        <td align="center"><b><%=count_8%>&nbsp;</b></td>
        <td align="center"><b><%=count_9%>&nbsp;</b></td>
        <td align="center"><b><%=count_10%>&nbsp;</b></td>
        <td align="center"><b><%=count_11%>&nbsp;</b></td>
        <td align="center"><b><%=count_12%>&nbsp;</b></td>
        <td align="center"><b><%=count_13%>&nbsp;</b></td>		
        <td align="center"><b><%=total%>&nbsp;</b></td>
	</tr>
</table>
</body>
</html>
<%
rsClient.Close();
%>