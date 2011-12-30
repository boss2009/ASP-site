<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan__intpLoan_Req_id = String(Request.QueryString("intLoan_Req_id"));
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_Loan_Request2("+ rsLoan__intpLoan_Req_id.replace(/'/g, "''") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();
%>

<html>
<head>
	<title><%=rsLoan.Fields.Item("chvLoan_name").Value%> - Loan Request No. <%=Request.QueryString("intLoan_Req_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="147,*" frameborder="0" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m008F3panel.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
  <frameset rows="24%,*" cols="*" frameborder="NO" border="0" framespacing="0" >
    <frame name="HeaderFrame" scrolling="NO"  src="m008F3Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m008FS01.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsLoan.Close();
%>