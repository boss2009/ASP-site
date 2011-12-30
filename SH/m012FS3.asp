<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#Include File="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2("+Request.QueryString("insSchool_id")+",'',0,0,0,0,0,0,0,'',1,'Q',0)}"
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();
%>
<html>
<head>
	<title><%=(rsInstitution.Fields.Item("chvSchool_Name").Value)%>- Institution ID: <%=Request.QueryString("insSchool_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="130,*" frameborder="0" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m012F3panel.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>">
  <frameset rows="12%,88%" cols="*" frameborder="NO" border="0" framespacing="0">
    <frame name="HeaderFrame" scrolling="NO" src="m012F3Hdr.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m012FS01.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsInstitution.Close();
%>