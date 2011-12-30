<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#Include File="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_contacts("+Request.QueryString("intContact_id")+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();
%>
<html>
<head>
	<title><%=rsContact.Fields.Item("chvFst_name").value%>&nbsp;<%=rsContact.Fields.Item("chvLst_name").value%> - Contact ID: <%=Request.QueryString("intContact_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="130,*" frameborder="0" framespacing="0">
	<frame name="MenuFrame" scrolling="NO" src="m004F3panel.asp?intContact_id=<%=Request.QueryString("intContact_id")%>">
	<frameset rows="20%,*" cols="*" frameborder="NO" border="0" framespacing="0">
		<frame name="HeaderFrame" scrolling="NO" src="m004F3Hdr.asp?intContact_id=<%=Request.QueryString("intContact_id")%>">
		<frame name="BodyFrame" scrolling="YES" src="m004e0101.asp?intContact_id=<%=Request.QueryString("intContact_id")%>">
	</frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsContact.Close();
%>