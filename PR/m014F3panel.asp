<!--------------------------------------------------------------------------
* File Name: m014F3Panel.asp
* Title: Purchase Requisition Panel
* Main SP: cp_frmpanel
* Description: This page displays the subsections of purchase requistion
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(14)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();

var rsRequisition = Server.CreateObject("ADODB.Recordset");
rsRequisition.ActiveConnection = MM_cnnASP02_STRING;
rsRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition(0,0,'',1,"+ Request.QueryString("insPurchase_Req_id")+ ",0)}";
rsRequisition.CursorType = 0;
rsRequisition.CursorLocation = 2;
rsRequisition.LockType = 3;
rsRequisition.Open();
%>
<html>
<head>
	<title>Purchase Requisition Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="JavaScript">
	<!--
	function MM_reloadPage(init) {  //reloads the window if Nav4 resized
		if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
			document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
		else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
	}
	MM_reloadPage(true);
	// -->
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="window.focus();">
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><div align="center"><a href="javascript: self.close();" target="_top"><img src="../i/tn_pur_req_02.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014e0101.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">General Information</a></td>
    </tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014q0201.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">Equipment Requested</a></td>
    </tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014q0301.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">Equipment Received</a></td>
    </tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014q0401.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">Requisition Notes</a></td>
    </tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014q0501.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">Backorder Received</a></td>
    </tr>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014e0602.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="BodyFrame">Forms & Reports</a></td>
    </tr>
<%
if ((rsRequisition.Fields.Item("insPurchase_sts_id").Value==6) || (rsRequisition.Fields.Item("insPurchase_sts_id").Value==7)) {
%>
	<tr>		
		<td height="18px" class="MenuItem" align="center"><a href="m014e0701.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="_blank">Create Inventory</a></td>
    </tr>
<%
}
%>
	<tr>
		<td height="18px" class="MenuItem">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m014a01j.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','wj014');" accesskey=D>Copy to DeskTop</a></td>
	</tr>
</table>
</body>
</html>
<%
rsFunction.Close();
%>