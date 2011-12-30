<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsEquipmentSupplied = Server.CreateObject("ADODB.Recordset");
rsEquipmentSupplied.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentSupplied.Source = "{call dbo.cp_company_equipment("+Request.QueryString("intEqCls_Dtl_id")+",0,'D',0)}";
rsEquipmentSupplied.CursorType = 0;
rsEquipmentSupplied.CursorLocation = 2;
rsEquipmentSupplied.LockType = 3;
rsEquipmentSupplied.Open();

Response.Redirect("AddDeleteSuccessful.asp?action=removed");
%>
<html>
<head>
	<title>Remove Equipment</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
</body>
</html>
