<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
switch (String(Request.Form("LinkToClass"))){
	//client
	case "1":
		var rsLinkContact = Server.CreateObject("ADODB.Recordset");
		rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
		rsLinkContact.Source="{call dbo.cp_clnctact2("+Request.Form("LinkToObject")+","+Request.QueryString("intContact_id")+","+Request.Form("Relationship")+","+Request.QueryString("KeyContact")+",0,'A',0)}";
		rsLinkContact.CursorType = 0;
		rsLinkContact.CursorLocation = 2;
		rsLinkContact.LockType = 3;
		rsLinkContact.Open();
	break;
	//company
	case "2":
		var rsLinkContact = Server.CreateObject("ADODB.Recordset");
		rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
		rsLinkContact.Source="{call dbo.cp_company_contact("+Request.Form("LinkToObject")+","+Request.Form("WorkType")+","+Request.QueryString("intContact_id")+",'A',0)}";
		rsLinkContact.CursorType = 0;
		rsLinkContact.CursorLocation = 2;
		rsLinkContact.LockType = 3;
		rsLinkContact.Open();
	break;
	//institution
	case "3":
		var rsLinkContact = Server.CreateObject("ADODB.Recordset");
		rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
		rsLinkContact.Source="{call dbo.cp_school_Contacts("+Request.Form("LinkToObject")+","+Request.QueryString("intContact_id")+"," +Request.Form("Relationship")+",1,'A',0)}";
		rsLinkContact.CursorType = 0;
		rsLinkContact.CursorLocation = 2;
		rsLinkContact.LockType = 3;
		rsLinkContact.Open();
	break;
	//on-site support
	case "4":
		var rsLinkContact = Server.CreateObject("ADODB.Recordset");
		rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
		rsLinkContact.Source="{call dbo.cp_pilat_site_support("+Request.Form("LinkToObject")+","+Request.QueryString("intContact_id")+",1,'A',0)}";
		rsLinkContact.CursorType = 0;
		rsLinkContact.CursorLocation = 2;
		rsLinkContact.LockType = 3;
		rsLinkContact.Open();
	break;
}
Response.Redirect("InsertSuccessful.html");
%>
<html>
<head>
	<title>Link Contact</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
</body>
</html>