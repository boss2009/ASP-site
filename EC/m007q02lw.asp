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
	<Script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=520,height=250,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	
	function EditClass(id, type){
		switch (type){
			case 'A':
				openWindow('m007e0101.asp?ClassID='+id,'EditAbstractClass');
			break;
			case 'S':
				openWindow('m007e0102.asp?ClassID='+id,'EditSubAbstractClass');
			break;
			case 'C':
				openWindow('m007e0103.asp?ClassID='+id,'EditConcreteClass');
			break;		
		}
	}
	
	function AddClass(type){
		switch (type){
			case 'S':
				openWindow('m007a0102.asp?ClassID=<%=Request.QueryString("ClassID")%>','NewSubAbstractClass');
			break;
		}
	}
	</Script>
</head>
<body onLoad="first.focus();">
<h3>Equipment Class - Hierarchy</h3>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
<a id=first class="blue" href="m007q01lw.asp">.. Back To Root</a><br>
<%
while (!rsSubAbstractClass.EOF){ 
%>
	<a href="m007q03lw.asp?ClassID=<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%>&ParentID=<%=rsSubAbstractClass.Fields.Item("insSuper_Class_id").Value%>"><img src="../i/collapse.gif" align="absmiddle" ALT="Expand SubAbstract Class <%=rsSubAbstractClass.Fields.Item("chvName").Value%>"></a><a class="green" href="javascript: EditClass('<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%>','S');"><%=rsSubAbstractClass.Fields.Item("chvName").Value%></a><br>
<%
	rsSubAbstractClass.MoveNext();
}
%>
<a href="javascript: AddClass('S');">Add Sub Abstract Class Under This Class</a>	
</body>
</html>
<%
rsSubAbstractClass.Close();
%>