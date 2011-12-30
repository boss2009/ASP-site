<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsReferrals = Server.CreateObject("ADODB.Recordset");
rsReferrals.ActiveConnection = MM_cnnASP02_STRING;
rsReferrals.Source = "{call dbo.cp_Referrals2("+ Request.QueryString("intAdult_id") + ",0,0,'',0,0,0,0,0,0,0,0,0,4,'Q',0)}";
rsReferrals.CursorType = 0;
rsReferrals.CursorLocation = 2;
rsReferrals.LockType = 3;
rsReferrals.Open();%>
<html>
<head>
	<title>Referral Type</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Referral Type</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="180">Date</th>	
		<th class="headrow" nowrap align="left">Type</th>
		<th class="headrow" nowrap align="left">Detail</th>
    </tr>
<% 
while (!rsReferrals.EOF) { 
%>
    <tr> 
		<td nowrap><%=FilterDate(rsReferrals.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
		<td nowrap><%=(rsReferrals.Fields.Item("chvType").Value)%>&nbsp;</td>
		<td nowrap><%=(rsReferrals.Fields.Item("chvDetails").Value)%>&nbsp;</td>
    </tr>
<%
	rsReferrals.MoveNext();
}
%>
</table>
<br><br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsReferrals.Close();
%>