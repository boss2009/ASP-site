<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_buyout_request3("+Request.QueryString("intBuyout_Req_id")+",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var rsBuyoutHeader = Server.CreateObject("ADODB.Recordset");
rsBuyoutHeader.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutHeader.Source = "{call dbo.cp_FrmHdr_10("+rsBuyout.Fields.Item("intEq_user_id").Value+","+rsBuyout.Fields.Item("insEq_user_type").Value+")}";
rsBuyoutHeader.CursorType = 0;
rsBuyoutHeader.CursorLocation = 2;
rsBuyoutHeader.LockType = 3;
rsBuyoutHeader.Open();
%>
<html>
<head>
	<title>Buyout Request Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px"> 
<%
if ((rsBuyoutHeader.EOF) || (rsBuyout.EOF)) {
%>
<i>Information not available for this Buyout.</i> <br>
<br>
<br>
<br>
<br>
<br>
<%
} else {
	switch (rsBuyout.Fields.Item("insEq_user_type").Value) {
		case 3:
%>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<td nowrap><b>Buyer Name:</b></td>
		<td width="170" nowrap><%=(rsBuyoutHeader.Fields.Item("chvBuyer_Name").Value)%></td>
		<td nowrap><b>Disability:</b></td>
		<td nowrap><%=(rsBuyoutHeader.Fields.Item("chvDisability").Value)%></td>
    </tr>
    <tr> 
		<td nowrap><b>User Type:</b></td>
		<td nowrap>Client</td>
		<td nowrap><b>Case Manager:</b></td>
		<td nowrap><%=(rsBuyoutHeader.Fields.Item("chvCase_Manager").Value)%></td>
    </tr>
    <tr>
		<td valign="top"><b>Address:</b></td>
		<td valign="top"> <%=(rsBuyoutHeader.Fields.Item("chvAddress").Value)%><br>
        	<%=(rsBuyoutHeader.Fields.Item("chvCity").Value)%>&nbsp; <%=(rsBuyoutHeader.Fields.Item("chrprvst_abbv").Value)%>&nbsp; 
	        <%=FormatPostalCode(rsBuyoutHeader.Fields.Item("chvPostal_zip").Value)%></td>
		<td></td>
		<td></td>
<!--
      <td valign="top"><b>Referral Type:</b></td>
      <td valign="top"><%=(rsBuyoutHeader.Fields.Item("chvReferral_Type").Value)%></td>
-->
    </tr>
    <tr> 
		<td nowrap valign="top"><b>Phone Number:</b></td>
		<td valign="top"><%=FormatPhoneNumber(rsBuyoutHeader.Fields.Item("chvPhone_Type_1").Value,rsBuyoutHeader.Fields.Item("chvPhone1_Arcd").Value,rsBuyoutHeader.Fields.Item("chvPhone1_Num").Value,rsBuyoutHeader.Fields.Item("chvPhone1_Ext").Value,rsBuyoutHeader.Fields.Item("chvPhone_Type_2").Value,rsBuyoutHeader.Fields.Item("chvPhone2_Arcd").Value,rsBuyoutHeader.Fields.Item("chvPhone2_Num").Value,rsBuyoutHeader.Fields.Item("chvPhone2_Ext").Value,"","","","")%></td>
		<td nowrap valign="top"><b>Referring Agent:</b></td>
		<td valign="top"><%=(rsBuyoutHeader.Fields.Item("chvReferring_Agent").Value)%></td>
	</tr>
</table>
<%
		break;
		case 4:
%>
<table cellspacing="1" cellpadding="1" border="0">
    <tr> 
		<td valign="top" nowrap><b>Buyer Name:</b></td>
		<td valign="top" width="180"><%=(rsBuyoutHeader.Fields.Item("chvBuyer_Name").Value)%></td>
		<td valign="top" nowrap><b>Referral Date:</b></td>
		<td valign="top"><%=FilterDate(rsBuyoutHeader.Fields.Item("dtsMost_Recent_Ref_Date").Value)%></td>
    </tr>
    <tr> 
		<td valign="top" nowrap><b>Referring Agent:</b></td>
		<td valign="top"><%=(rsBuyoutHeader.Fields.Item("chvReferring_Agent").Value)%></td>
		<td valign="top" nowrap><b>Case Manager:</b></td>
		<td valign="top"><%=(rsBuyoutHeader.Fields.Item("chvCase_Manager").Value)%></td>
    </tr>
    <tr> 
		<td valign="top" nowrap><b>Address:</b></td>
		<td valign="top" width="220"> <%=(rsBuyoutHeader.Fields.Item("chvAddress").Value)%><br>
			<%=(rsBuyoutHeader.Fields.Item("chvCity").Value)%>&nbsp; <%=(rsBuyoutHeader.Fields.Item("chrprvst_abbv").Value)%>&nbsp; 
			<%=FormatPostalCode(rsBuyoutHeader.Fields.Item("chvPostal_zip").Value)%> </td>
		<td></td>
		<td></td>
<!--		
      <td valign="top" nowrap><b>Referral Type:</b></td>
      <td valign="top"><%=(rsBuyoutHeader.Fields.Item("chvReferral_Type").Value)%></td>
-->
	</tr>
</table>
<%
		break;
	}
}
%>
</div>
</body>
</html>
<%
rsBuyout.Close();
rsBuyoutHeader.Close();
%>