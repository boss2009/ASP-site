<!--------------------------------------------------------------------------
* File Name: m014e0602.asp
* Title: Purchase Requisition Forms & Reports
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JavaScript"%>
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

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
	<title>Forms & Reports</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Forms & Reports</h5>
<hr>
<a href="RequisitionForm.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="blank">Purchase Requisition Form</a><br>
<%
if (rsRequisition.Fields.Item("insPurchase_sts_id").Value==6) {
%>
<a href="ReceivingReport.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="blank">Purchase Receiving Report</a><br>
<%
}
%>
<a href="Fax.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>" target="blank">Facsmile Transmission</a><br>
</body>
</html>