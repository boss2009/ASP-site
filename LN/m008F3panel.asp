<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(8)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Loan Request Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="JavaScript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}	
</script>
</head>
<body onLoad="window.focus();">
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_loan_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
<% 
while (!rsFunction.EOF) { 
%>
    <tr> 
		<td height="18px" class="MenuItem" align="center"><a href="<%=(rsFunction.Fields.Item("filename").Value)%>?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="BodyFrame"><%=(rsFunction.Fields.Item("name").Value)%></a></td>
    </tr>
<%
	rsFunction.MoveNext();
}
%>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m008q1001.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="BodyFrame">Forms and Reports</a></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m008q1101.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>" target="BodyFrame">Correspondence</a></td>
	</tr>	
<!--
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m008q0301b.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>" target="BodyFrame">Test Equip Loaned.</a></td>
	</tr>	
-->
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m008a01j.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>','wj0101');" accesskey=D>Copy to DeskTop</a></td>
	</tr>	
</table>
</body>
</html>
<%
rsFunction.Close();
%>