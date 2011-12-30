<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsClientAddress = Server.CreateObject("ADODB.Recordset");
rsClientAddress.ActiveConnection = MM_cnnASP02_STRING;
rsClientAddress.Source = "{call dbo.cp_Adult_Address("+ Request.QueryString("intAdult_id") + ")}";
rsClientAddress.CursorType = 0;
rsClientAddress.CursorLocation = 2;
rsClientAddress.LockType = 3;
rsClientAddress.Open();

var ChkAdultAddress = Server.CreateObject("ADODB.Command");
ChkAdultAddress.ActiveConnection = MM_cnnASP02_STRING;
ChkAdultAddress.CommandText = "dbo.cp_Chk_Adult_Address";
ChkAdultAddress.CommandType = 4;
ChkAdultAddress.CommandTimeout = 0;
ChkAdultAddress.Prepared = true;
ChkAdultAddress.Parameters.Append(ChkAdultAddress.CreateParameter("RETURN_VALUE", 3, 4));
ChkAdultAddress.Parameters.Append(ChkAdultAddress.CreateParameter("@intAdult_id", 3, 1,10000,Request.QueryString("intAdult_id")));
ChkAdultAddress.Parameters.Append(ChkAdultAddress.CreateParameter("@insRtnFlag", 2, 2));
ChkAdultAddress.Execute();
%>
<html>
<head>
	<title>Client Address</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();	   
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=500,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body>
<h5>Client Address</h5>
<i>Please note: all correspondence and equipment will be shipped to the "Address While Attending School".</i>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th nowrap class="headrow" align="left">Type</th>
		<th nowrap class="headrow" align="left" width="180">Address</th>
		<th nowrap class="headrow" align="left" width="120">City</th>
		<th nowrap class="headrow" align="left">Province</th>
		<th nowrap class="headrow" align="left">Postal Code</th>
		<th nowrap class="headrow" align="left">Phone Number</th>
		<th nowrap class="headrow" align="left">E-mail</th>
	</tr>
<% 
while (!rsClientAddress.EOF) { 
%>
    <tr> 
		<td valign="top" nowrap><a href="m001e0301.asp?intaddr_id=<%=(rsClientAddress.Fields.Item("intaddr_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>"><%=((rsClientAddress.Fields.Item("chvAddrs_type").Value=="P")?"Permanent":"Address While Attending School")%></a></td>
		<td valign="top"><%=(rsClientAddress.Fields.Item("chvAddress").Value)%>&nbsp;</td>
		<td valign="top"><%=(rsClientAddress.Fields.Item("chvCity").Value)%>&nbsp;</td>
		<td valign="top" align="center"><%=(rsClientAddress.Fields.Item("chvProv").Value)%>&nbsp;</td>
		<td valign="top"><%=FormatPostalCode(rsClientAddress.Fields.Item("chvPostal_zip").Value)%>&nbsp;</td>
		<td valign="top" nowrap><%=FormatPhoneNumber(rsClientAddress.Fields.Item("chvPhone_Type1").Value,rsClientAddress.Fields.Item("chvPhone1_Arcd").Value,rsClientAddress.Fields.Item("chvPhone1_Num").Value,rsClientAddress.Fields.Item("chvPhone1_Ext").Value,rsClientAddress.Fields.Item("chvPhone_Type2").Value,rsClientAddress.Fields.Item("chvPhone2_Arcd").Value,rsClientAddress.Fields.Item("chvPhone2_Num").Value,rsClientAddress.Fields.Item("chvPhone2_Ext").Value,"","","","")%>&nbsp;</td>
		<td valign="top"><%=(rsClientAddress.Fields.Item("chvemail").Value)%>&nbsp;</td>
    </tr>
<%
	rsClientAddress.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
<%
switch (ChkAdultAddress.Parameters.Item("@insRtnFlag").Value) {
	case 0:
%>
		<td nowrap width="170"><a href="javascript: openWindow('m001a0301.asp?AddressType=Permanent&intAdult_id=<%=Request.QueryString("intAdult_id")%>','wA0301');">Add Permanent Address</a></td>
		<td nowrap><a href="javascript: openWindow('m001a0301.asp?AddressType=Address While Attending School&intAdult_id=<%=Request.QueryString("intAdult_id")%>','wA0301');">Add Address While Attending School</a></td>
<%
	break ;
	case 1:
%>
		<td><a href="javascript: openWindow('m001a0301.asp?AddressType=Address While Attending School&intAdult_id=<%=Request.QueryString("intAdult_id")%>','wA0301');">Add Address While Attending School</a></td>
<%
	break ;
	case 2:
%>
		<td><a href="javascript: openWindow('m001a0301.asp?AddressType=Permanent&intAdult_id=<%=Request.QueryString("intAdult_id")%>','wA0301');">Add Permanent Address</a></td>
<%
	break ;
}
%>
	</tr>
</table>
</body>
</html>
<%
rsClientAddress.Close();
%>