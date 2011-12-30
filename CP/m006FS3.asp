<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#Include File="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_company2("+Request.QueryString("intCompany_id")+",'',0,0,0,0,0,1,0,'',1,'Q',0)}"
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();	
%>
<html>
<head>
	<title><%=rsCompany.Fields.Item("chvCompany_Name").Value%> - <%=rsCompany.Fields.Item("chvWork_type_desc").Value%> ID: <%=Request.QueryString("intCompany_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="130,*" frameborder="0" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m006F3panel.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>">
  <frame name="BodyFrame" scrolling="YES" src="m006e0101.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>">
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsCompany.Close();
%>