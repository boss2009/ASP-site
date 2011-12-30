<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsAttachment = Server.CreateObject("ADODB.Recordset");
rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
rsAttachment.Source = "{call dbo.cp_get_loan_accessory2("+Request.QueryString("intLoan_Req_id")+",0,0)}";
rsAttachment.CursorType = 0;
rsAttachment.CursorLocation = 2;
rsAttachment.LockType = 3;
rsAttachment.Open();
%>
<html>
<head>
	<title>Attachments</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Attachments</h5>
<hr>
<table cellspacing="1" cellpadding="1">
	<tr>
		<th class="headrow" nowrap align="left" width="200">Accessory</th>	
		<th class="headrow" nowrap align="left">Quantity</th>
    </tr>
<% 
while (!rsAttachment.EOF) { 
%>
    <tr> 
		<td align="left"><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%></td>
		<td align="center"><input type="text" name="Quantity" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" readonly></td>
    </tr>
<%
	rsAttachment.MoveNext();
}
%>
</table>
<hr>
</body>
</html>
<%
rsAttachment.Close();
%>