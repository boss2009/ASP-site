<!--------------------------------------------------------------------------
* File Name: m012q03lw.asp
* Title: Concrete Classes
* Main SP: cp_Eqp_Class_LW
* Description: This page lists all the concrete classes
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW("+Request.QueryString("ClassID")+",'P',3)}";
rsConcreteClass.CursorType = 0;
rsConcreteClass.CursorLocation = 2;
rsConcreteClass.LockType = 3;
rsConcreteClass.Open();
%>
<html>
<head>
	<title>Concrete Classes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	function SelectClass(id, name){	
		if (!top.opener.closed) {	
	<%
	if (String(Request.QueryString("intEqpRequest_id")) == "") { 
	%>
			top.opener.document.frm0203.action = "m012a0203.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&ClassID="+id+"&ClassName="+name;
	<% 
	} else { 
	%>
			top.opener.document.frm0203.action = "m012e0203.asp?intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>&intReferral_id=<%=Request.QueryString("intReferral_id")%>&ClassID="+id+"&ClassName="+name;
	<% 
	} 
	%>		
			top.opener.document.frm0203.submit();
		}
		top.window.close();
	}	
	</Script>
</head>
<body>
<h5>Inventory Classes</h5>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
<a id=first class="green" href="m012q02lw.asp?ClassID=<%=Request.QueryString("ParentID")%>&intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>&intReferral_id=<%=Request.QueryString("intReferral_id")%>">.. Back To Parent Class</a><br>
<%
while (!rsConcreteClass.EOF){ 
%>
	<img src="../i/leaf.gif" align="absmiddle" ALT="Leaf Concrete Class <%=rsConcreteClass.Fields.Item("chvName").Value%>"><a class="red" href="javascript: SelectClass('<%=rsConcreteClass.Fields.Item("insAEquip_Class_id").Value%>','<%=FilterQuotes(rsConcreteClass.Fields.Item("chvName").Value)%>');"><%=rsConcreteClass.Fields.Item("chvName").Value%></a><br>
<%
	rsConcreteClass.MoveNext();
}
%>
</body>
</html>
<%
rsConcreteClass.Close();
%>