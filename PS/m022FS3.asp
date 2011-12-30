<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<%
var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStudent.Source = "{call dbo.cp_pilat_student("+Request.QueryString("intPStdnt_id")+",'','','','','',0,0,0,0,'',0,0,0,1,0,'',1,'Q',0)}";
rsPILATStudent.CursorType = 0;
rsPILATStudent.CursorLocation = 2;
rsPILATStudent.LockType = 3;
rsPILATStudent.Open();
%>

<html>
<head>
	<title><%=rsPILATStudent.Fields.Item("chvFst_name").Value%>&nbsp;<%=rsPILATStudent.Fields.Item("chvLst_name").Value%> - Temp Student No. <%=Request.QueryString("intPStdnt_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="147,*" frameborder="0" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m022F3panel.asp?intPStdnt_id=<%=Request.QueryString("intPStdnt_id")%>">
  <frameset rows="20%,80%" cols="*" frameborder="NO" border="0" framespacing="0" >
    <frame name="HeaderFrame" scrolling="NO"  src="m022F3Hdr.asp?intPStdnt_id=<%=Request.QueryString("intPStdnt_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m022e0101.asp?intPStdnt_id=<%=Request.QueryString("intPStdnt_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsPILATStudent.Close();
%>