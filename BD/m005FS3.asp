<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#Include File="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsBundle = Server.CreateObject("ADODB.Recordset");
rsBundle.ActiveConnection = MM_cnnASP02_STRING;
rsBundle.Source = "{call dbo.cp_Bundle("+Request.QueryString("insBundle_id")+",'',0.0,0,1,1,'',0,"+Session("insStaff_id")+",0,0,'',1,'Q',0)}"
rsBundle.CursorType = 0;
rsBundle.CursorLocation = 2;
rsBundle.LockType = 3;
rsBundle.Open();	
%>
<html>
<head>
	<title><%=rsBundle.Fields.Item("chvName").Value%> - Bundle ID: <%=Request.QueryString("insBundle_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="155,*" cols="*" frameborder="0" framespacing="0">
  <frame name="GeneralInfoFrame" scrolling="NO" src="m005e0101.asp?insBundle_id=<%=Request.QueryString("insBundle_id")%>">
  <frame name="ComponentFrame" scrolling="NO" src="m005e0102.asp?insBundle_id=<%=Request.QueryString("insBundle_id")%>">
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsBundle.Close();
%>