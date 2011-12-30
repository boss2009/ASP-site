<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#Include File="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_frmhdr(2,"+Request.QueryString("insStaff_id")+")}"
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title><%=rsStaff.Fields.Item("chvFst_name").Value%>&nbsp;<%=rsStaff.Fields.Item("chvLst_name").Value%> - Staff ID: <%=Request.QueryString("insStaff_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="130,*" frameborder="0" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m002F3panel.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>">
  <frameset rows="20%,80%" cols="*" frameborder="NO" border="0" framespacing="0">
    <frame name="HeaderFrame" scrolling="NO" src="m002F3Hdr.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m002FS01.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsStaff.Close();
%>