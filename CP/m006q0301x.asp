<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
// set the form action variable
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsCompanyContact = Server.CreateObject("ADODB.Recordset");
rsCompanyContact.ActiveConnection = MM_cnnASP02_STRING;
rsCompanyContact.Source = "{call dbo.cp_company_contact(0,0,"+Request.QueryString("intContact_id")+",'D',0)}";
rsCompanyContact.CursorType = 0;
rsCompanyContact.CursorLocation = 2;
rsCompanyContact.LockType = 3;
rsCompanyContact.Open();

Response.Redirect("AddDeleteSuccessful.asp?action=removed");
%>
<html>
<head>
	<title>Remove Contact</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
</body>
</html>
