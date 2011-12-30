<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ClassID") + ",'C',0)}";
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
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=700,height=500,scrollbars=1,left=0,top=0,status=1");
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
				openWindow('m007FS3.asp?ClassID='+id,'EditConcreteClass');
			break;		
		}
	}	
	
	function AddClass(type){
		switch (type){
			case 'C':
				openWindow('m007a0103.asp?ParentID=<%=Request.QueryString("ClassID")%>','NewConcreteClass');
			break;		
		}
	}
	</Script>
</head>
<body onLoad="first.focus();">
<h3>Equipment Class - Hierarchy</h3>
<span class="blue">Abstract Class</span> | <span class="green">SubAbstract Class</span> | <span class="red">Concrete Class</span> | <span class="grey">Inactive Class</span><br>
<hr>
	<a id=first class="green" href="m007q02lw.asp?ClassID=<%=Request.QueryString("ParentID")%>">.. Back To Parent Class</a><br>
	<%
	while (!rsConcreteClass.EOF){ 
	%>
		<img src="../i/leaf.gif" align="absmiddle" ALT="Leaf Concrete Class <%=rsConcreteClass.Fields.Item("chvName").Value%> <%=((rsConcreteClass.Fields.Item("bitIs_Class_Active").Value=="0")?" (Inactive)":"")%>"><a class=" <%=((rsConcreteClass.Fields.Item("bitIs_Class_Active").Value=="0")?"grey":"red")%>" href="javascript: EditClass('<%=rsConcreteClass.Fields.Item("insEquip_Class_id").Value%>','C');"><%=rsConcreteClass.Fields.Item("chvName").Value%></a><br>
	<%
		rsConcreteClass.MoveNext();
	}
	%>
	<a href="javascript: AddClass('C');">Add Concrete Class Under This Class</a>
</body>
</html>
<%
rsConcreteClass.Close();
%>