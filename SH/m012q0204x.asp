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

var rsOnsiteSupport = Server.CreateObject("ADODB.Recordset");
rsOnsiteSupport.ActiveConnection = MM_cnnASP02_STRING;
rsOnsiteSupport.Source = "{call dbo.cp_pilat_site_support("+Request.QueryString("intReferral_id")+","+Request.QueryString("intContact_id")+",0,'D',0)}";
rsOnsiteSupport.CursorType = 0;
rsOnsiteSupport.CursorLocation = 2;
rsOnsiteSupport.LockType = 3;
rsOnsiteSupport.Open();
Response.Redirect("AddDeleteSuccessful.asp?action=removed");
%>
<html>
<head>
	<title>Remove On-Site Support</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
</body>
</html>
