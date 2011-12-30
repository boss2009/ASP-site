<%@language="JAVASCRIPT"%>

<!--#include file="../inc/ASPUtility.inc" -->

<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<%
var rsClient__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client_Detail("+ rsClient__intpAdult_id.replace(/'/g, "''") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>

<html>
<head>
	<title><%=rsClient.Fields.Item("chvName").Value%> - Client No. <%=Request.QueryString("intAdult_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="147,*" frameborder="No" framespacing="0">
  <frame name="MenuFrame" scrolling="NO" src="m001F3panel.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
  <frameset rows="20%,80%" cols="*" frameborder="No" framespacing="0" >
    <frame name="HeaderFrame" scrolling="NO"  src="m001F3Hdr.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m001FS01.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsClient.Close();
%>