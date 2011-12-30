<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc"-->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ClassID") + ",'C',1)}";	
rsConcreteClass.CursorType = 0;
rsConcreteClass.CursorLocation = 2;
rsConcreteClass.LockType = 3;
rsConcreteClass.Open();	
%>
<html>
<head>
	<title>Equipment Class: <%=rsConcreteClass.Fields.Item("chvName").Value%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="135,*" frameborder="0" framespacing="0">	
	<frame name="EquipmentClassFrameHeader" scrolling="NO" noresize src="m007FS3Panel.asp?<%=Request.QueryString%>">
	<frame name="EquipmentClassFrameBody" src="m007e0103.asp?<%=Request.QueryString%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>