<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(0,'A',0)}";
rsAbstractClass.CursorType = 0;
rsAbstractClass.CursorLocation = 2;
rsAbstractClass.LockType = 3;
rsAbstractClass.Open();
%>
<html>
<head>
	<title>Abstract Classes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Inventory Classes</h5>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
<% 
while (!rsAbstractClass.EOF) {
%>	
	<a href="m008q02lw.asp?ClassID=<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>"><img src="../i/collapse.gif" align="absmiddle" ALT="Expand Abstract Class <%=rsAbstractClass.Fields.Item("chvName").Value%>"></a><%=rsAbstractClass.Fields.Item("chvName").Value%><br>
<%
	rsAbstractClass.MoveNext();
}
%>
</body>
</html>
<%
rsAbstractClass.Close();
%>