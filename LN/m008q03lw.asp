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
	function SelectClass(id, name, LUC){
		if (!top.opener.closed) {	
	<% 
	if (String(Request.QueryString("intEqpReq_Id")) == "") { 
	%>
			top.opener.document.frm08a0201.action = "m008a0201.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&ClassID="+id+"&ClassName="+name+"&LUC="+LUC;
	<% 
	} else { 
	%>
			top.opener.document.frm08a0201.action = "m008e0201.asp?intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&ClassID="+id+"&ClassName="+name+"&LUC="+LUC;
	<% 
	} 
	%>			
			top.opener.document.frm08a0201.submit();
		}
		top.window.close();
	}	
	</Script>
</head>
<body>
<h5>Inventory Classes</h5>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
<a id=first class="green" href="m008q02lw.asp?intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>&ClassID=<%=Request.QueryString("ParentID")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>">.. Back To Parent Class</a><br>
<%
while (!rsConcreteClass.EOF){ 
%>
	<img src="../i/leaf.gif" align="absmiddle" ALT="Leaf Concrete Class <%=rsConcreteClass.Fields.Item("chvName").Value%>"><a class="red" href="javascript: SelectClass('<%=rsConcreteClass.Fields.Item("insAEquip_Class_id").Value%>','<%=FilterQuotes(rsConcreteClass.Fields.Item("chvName").Value)%>','<%=(rsConcreteClass.Fields.Item("fltList_Unit_Cost").Value)%>');"><%=rsConcreteClass.Fields.Item("chvName").Value%></a><br>
<%
	rsConcreteClass.MoveNext();
}
%>
</body>
</html>
<%
rsConcreteClass.Close();
%>