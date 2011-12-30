<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();
%>
<html>
<head>
<title>Training Frame</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
switch (String(rsLoan.Fields.Item("insEq_user_type").Value)){
	//client
	case "3":		
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m008F06Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m008e0601.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
<%
	break;
	//institution
	case "4":
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m008F06Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m008e0601.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
<%	
	break;
	default:
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m008F06Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" >
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
<%
	break;
}
%>		
</html>

