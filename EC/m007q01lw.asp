<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
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
		}
	}
	
	function AddClass(type){
		switch (type){
			case 'A':
				openWindow('m007a0101.asp','NewAbstractClass');
			break;
		}
	}
	</Script>
</head>
<body onLoad="window.focus();">
<h3>Equipment Class - Hierarchy</h3>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span><br>
<hr>
	<% 
	while (!rsAbstractClass.EOF) {
	%>	
		<a href="m007q02lw.asp?ClassID=<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>"><img src="../i/collapse.gif" align="absmiddle" ALT="Expand Abstract Class <%=rsAbstractClass.Fields.Item("chvName").Value%>"></a>
		<a class="blue" href="javascript: EditClass('<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>','A');"><%=rsAbstractClass.Fields.Item("chvName").Value%></a>
		<br>
	<%
		rsAbstractClass.MoveNext();
	}
	%>
	<a href="javascript: AddClass('A')">Add An Abstract Class</a>
</body>
</html>
<%
rsAbstractClass.Close();
%>