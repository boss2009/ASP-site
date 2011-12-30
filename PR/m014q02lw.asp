<!--------------------------------------------------------------------------
* File Name: m014q02lw.asp
* Title: Sub Abstract Classes
* Main SP: cp_Eqp_Class_LW
* Description: This page lists all the subabstract classes.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW("+Request.QueryString("ClassID")+",'S',0)}";
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();
%>
<html>
<head>
	<title>Sub Abstract Classes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Inventory Classes</h5>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
<a id=first class="blue" href="m014q01lw.asp?insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>&insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>">.. Back To Root</a><br>
<%
while (!rsSubAbstractClass.EOF){ 
%>
	<a href="m014q03lw.asp?ClassID=<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%>&ParentID=<%=rsSubAbstractClass.Fields.Item("insSuper_Class_id").Value%>&insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>"><img src="../i/collapse.gif" align="absmiddle" ALT="Expand SubAbstract Class <%=rsSubAbstractClass.Fields.Item("chvName").Value%>"></a><%=rsSubAbstractClass.Fields.Item("chvName").Value%><br>
<%
	rsSubAbstractClass.MoveNext();
}
%>
</body>
</html>
<%
rsSubAbstractClass.Close();
%>