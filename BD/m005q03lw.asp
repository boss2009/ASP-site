<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW("+Request.QueryString("ClassID")+",'C',0)}";
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
	<script language="Javascript">
	function SelectClass(id, name){
		if (!top.opener.closed) {	
			top.opener.document.frm05s01.ClassName.value=name;
			top.opener.document.frm05s01.ClassID.value=id;
		}
		top.window.close();
	}
	</script>
</head>
<body>
<h5>Inventory Classes</h5>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
	<a id=first class="green" href="m005q02lw.asp?ClassID=<%=Request.QueryString("ParentID")%>">.. Back To Parent Class</a><br>
	<%
	while (!rsConcreteClass.EOF){ 
	%>
		<img src="../i/leaf.gif" align="absmiddle" ALT="Leaf Concrete Class <%=rsConcreteClass.Fields.Item("chvName").Value%>"><a class="red" href="javascript: SelectClass('<%=rsConcreteClass.Fields.Item("insEquip_Class_id").Value%>','<%=FilterQuotes(rsConcreteClass.Fields.Item("chvName").Value)%>','C');"><%=rsConcreteClass.Fields.Item("chvName").Value%></a><br>
	<%
		rsConcreteClass.MoveNext();
	}
	%>
</body>
</html>
<%
rsConcreteClass.Close();
%>