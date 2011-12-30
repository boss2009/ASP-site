<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
//Generic Loans and Buyouts are suppressed.

var cmdFundingSource = Server.CreateObject("ADODB.Command");
cmdFundingSource.ActiveConnection = MM_cnnASP02_STRING;
cmdFundingSource.CommandText = "dbo.cp_Get_EqpSrv_FundingSrc";
cmdFundingSource.CommandType = 4;
cmdFundingSource.CommandTimeout = 0;
cmdFundingSource.Prepared = true;
cmdFundingSource.Parameters.Append(cmdFundingSource.CreateParameter("RETURN_VALUE", 3, 4));
cmdFundingSource.Parameters.Append(cmdFundingSource.CreateParameter("@intEquip_set_id", 3, 1,1,Request.QueryString("intEquip_Set_id")));
cmdFundingSource.Parameters.Append(cmdFundingSource.CreateParameter("@insRtnFlag", 2, 2));
cmdFundingSource.Execute();

if (!((cmdFundingSource.Parameters.Item("@insRtnFlag").Value == -1) || (cmdFundingSource.Parameters.Item("@insRtnFlag").Value == null))) {
	var rsFundingSource = Server.CreateObject("ADODB.Recordset");
	rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsFundingSource.Source = "{call dbo.cp_get_eqpsrv_fundingsrc("+ Request.QueryString("intEquip_Set_id") + ",0)}";
	rsFundingSource.CursorType = 0;
	rsFundingSource.CursorLocation = 2;
	rsFundingSource.LockType = 3;
	rsFundingSource.Open();
}
%>
<html>
<head>
	<title>Funding Source</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Funding Source</h5>
<hr>
<%
if ((cmdFundingSource.Parameters.Item("@insRtnFlag").Value == -1) || (cmdFundingSource.Parameters.Item("@insRtnFlag").Value == null)) {
%>
<i>This inventory is current not on loan or sold.  Funding Source not available.</i>
<%
} else {
%>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr>
		<th class="headrow" align="left">Referral Date</th>		
		<th class="headrow" align="left">Referral Type</th> 		
		<th class="headrow" align="left">Funding Source</th>
    </tr>
<% 
while (!rsFundingSource.EOF) {
	if ((rsFundingSource.Fields.Item("insRefAgt_id").Value != 23) && (Trim(rsFundingSource.Fields.Item("chvfunding_source_name").Value) != "Generic Loan")) {
%>
	<tr> 
		<td nowrap><%=FilterDate(rsFundingSource.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
		<td nowrap><%=Trim(rsFundingSource.Fields.Item("chvRefAgt").Value)%>&nbsp;</td>
		<td nowrap><%=Trim(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%>&nbsp;</td>		
	</tr>
<%
	}
	rsFundingSource.MoveNext();
}
rsFundingSource.Close();
%>
</table>
<%
}
%>
<br><br><br>
<input type="button" value="Close" onclick="window.close();" class="btnstyle">
</body>
</html>