<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client_Detail("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_Loan_Request_LW("+ Request.QueryString("intAdult_id") + ",1,0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_Request("+ Request.QueryString("intAdult_id") + ")}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();
%>
<html>
<head>
	<title>Summary for <%=(rsClient.Fields.Item("chvName").Value)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body onLoad="window.focus();">
<b><u>General Information</u></b>
<table cellpadding="1" cellspacing="3">
    <tr> 
		<td><%=(rsClient.Fields.Item("chvName").Value)%></td>
    </tr>
    <tr> 
		<td><%=FormatSIN(rsClient.Fields.Item("chrSIN_no").Value)%></td>
    </tr>
    <tr> 
		<td><%=FilterDate(rsClient.Fields.Item("dtsBirth_date").Value)%></td>
    </tr>
    <tr> 
		<td><%=(rsClient.Fields.Item("chvAddress").Value)%></td>
    </tr>
    <tr> 
		<td><%=(rsClient.Fields.Item("chvCity").Value)%></td>
    </tr>
    <tr> 
		<td><%=(rsClient.Fields.Item("chrprvst_abbv").Value)%></td>
    </tr>
    <tr> 
		<td><%=(rsClient.Fields.Item("chvcntry_name").Value)%></td>
    </tr>
    <tr> 
		<td><%=FormatPostalCode(rsClient.Fields.Item("chvPostal_zip").Value)%></td>
    </tr>
    <tr> 
		<td><%=FormatPhoneNumber(rsClient.Fields.Item("chvPhoneName").Value,rsClient.Fields.Item("chvPhone1_Arcd").Value,rsClient.Fields.Item("chvPhone1_Num").Value,rsClient.Fields.Item("chvPhone1_Ext").Value,"","","","","","","","")%></td>
    </tr>
</table>
<br><br>
<b><u>Equipment On Loan</u></b>
<table cellpadding="1" cellspacing="3" width="400">
	<tr>
		<td nowrap><b>Inventory Name</b></td>
		<td nowrap><b>Equipment Status</b></td>		
	</tr>
<%
while (!rsLoan.EOF) {
%>
	<tr>
		<td nowrap><%=rsLoan.Fields.Item("chvInventory_Name").Value%>&nbsp;</td>
		<td nowrap><%=rsLoan.Fields.Item("chvequip_status").Value%>&nbsp;</td>
	</tr>
<%
	rsLoan.MoveNext();
}
%>
</table>
<br><br>
<b><u>Buyouts</u></b>
<table cellpadding="1" cellspacing="3" width="600">
	<tr>
		<td nowrap><b>Inventory Name</b></td>
		<td nowrap><b>Equipment Status</b></td>
		<td nowrap><b>Sold Price</b></td>		
		<td nowrap><b>Payment Status</b></td>
	</tr>
<%
while (!rsBuyout.EOF) {
%>
	<tr>
		<td nowrap><%=rsBuyout.Fields.Item("chvInventory_Name").Value%>&nbsp;</td>
		<td nowrap align="center"><%=rsBuyout.Fields.Item("chvequip_status").Value%>&nbsp;</td>
		<td nowrap align="right"><%=FormatCurrency(rsBuyout.Fields.Item("fltEqp_Sold_price").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=rsBuyout.Fields.Item("chvPayStatus").Value%>&nbsp;</td>						
	</tr>
<%
	rsBuyout.MoveNext();
}
%>
</table>
<br><br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsClient.Close();
rsLoan.Close();
rsBuyout.Close();
%>