<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsReferrals = Server.CreateObject("ADODB.Recordset");
rsReferrals.ActiveConnection = MM_cnnASP02_STRING;
rsReferrals.Source = "{call dbo.cp_Referrals2("+ Request.QueryString("intEq_user_id") + ",0,0,'',0,0,0,0,0,0,0,0,0,4,'Q',0)}";
rsReferrals.CursorType = 0;
rsReferrals.CursorLocation = 2;
rsReferrals.LockType = 3;
rsReferrals.Open();
%>
<html>
<head>
	<title>Referral Type</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=700,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>
</head>
<body>
<h5>Referral Type</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr>
		<th class="headrow" align="left">Type</th>
		<th class="headrow" align="left">Date</th>		
		<th class="headrow" align="left">Details</th> 		
    </tr>
<% 
while (!rsReferrals.EOF) { 
%>
	<tr> 
		<td nowrap><%=(rsReferrals.Fields.Item("chvType").Value)%>&nbsp;</td>
		<td nowrap><%=FilterDate(rsReferrals.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
		<td nowrap><%=(rsReferrals.Fields.Item("chvDetails").Value)%>&nbsp;</td>
	</tr>
<%
	rsReferrals.MoveNext();
}
%>
</table>
<br><br><br>
<input type="button" value="Close" onclick="window.close();" class="btnstyle">
</body>
</html>
<%
rsReferrals.Close();
%>