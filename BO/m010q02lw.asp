<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsSubAbstractClasss = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClasss.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClasss.Source = "{call dbo.cp_Eqp_Class_LW("+Request.QueryString("ClassID")+",'S',0)}";
rsSubAbstractClasss.CursorType = 0;
rsSubAbstractClasss.CursorLocation = 2;
rsSubAbstractClasss.LockType = 3;
rsSubAbstractClasss.Open();
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
<a id=first class="blue" href="m010q01lw.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>">.. Back To Root</a><br>
<%
while (!rsSubAbstractClasss.EOF){ 
%>
	<a href="m010q03lw.asp?ClassID=<%=rsSubAbstractClasss.Fields.Item("insEquip_Class_id").Value%>&ParentID=<%=rsSubAbstractClasss.Fields.Item("insSuper_Class_id").Value%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>"><img src="../i/collapse.gif" align="absmiddle" ALT="Expand SubAbstract Class <%=rsSubAbstractClasss.Fields.Item("chvName").Value%>"></a><%=rsSubAbstractClasss.Fields.Item("chvName").Value%><br>
<%
	rsSubAbstractClasss.MoveNext();
}
%>
</body>
</html>
<%
rsSubAbstractClasss.Close();
%>