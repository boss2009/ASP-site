<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (Request.QueryString("MM_edit") == "true"){
	var rsAttachment = Server.CreateObject("ADODB.Recordset");
	rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
	rsAttachment.CursorType = 0;
	rsAttachment.CursorLocation = 2;
	rsAttachment.LockType = 3;	
	for (var i=1; i<=Request.QueryString("ArraySize"); i++){
		rsAttachment.Source = "{call dbo.cp_Update_EqpClss_Attachment(" + Request.QueryString("ClassID") + "," + Request.QueryString("Accessory")(i) + "," + Request.QueryString("Quantity")(i) + "," + Session("insStaff_id") + ",0)}";
		rsAttachment.Open();
		rsAttachment.Close();			
	}
	Response.Redirect("m007e0103b.asp?ClassID="+Request.QueryString("ClassID"));
}

var rsAttachment = Server.CreateObject("ADODB.Recordset");
rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
rsAttachment.Source = "{call dbo.cp_Get_EqpClss_Attachment("+ Request.QueryString("ClassID") + ",0)}";
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
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				document.frm0103b.submit();
			break;
	   		case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
</head>
<body>
<form name="frm0103b" method="GET" action="<%=MM_editAction%>">
<h5>Attachments</h5>
<a href="m007e0103b.asp?All=1&ClassID=<%=Request.QueryString("ClassID")%>">Show All</a> | <a href="m007e0103b.asp?All=0&ClassID=<%=Request.QueryString("ClassID")%>">Show Only Checked</a>
<hr>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<th class="headrow" width="200">Accessory</th>	
		<th class="headrow">Quantity</th>
    </tr>
<% 
var count = 0;
while (!rsAttachment.EOF) { 
	if (String(Request.QueryString("All"))=="1"){ 
%>
    <tr> 
		<td nowrap><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%></td>
		<td nowrap>
			<input type="text" name="Quantity" maxlength=5 size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" onKeypress="AllowNumericOnly();">
			<input type="hidden" name="Accessory" value="<%=(rsAttachment.Fields.Item("insAttachment_id").Value)%>">			
		</td>
    </tr>
<%	
		count++;
	} else {
		if (String(rsAttachment.Fields.Item("insQuantity").Value) != "0") { 
%>
    <tr>
		<td nowrap><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%></td>
		<td nowrap>
			<input type="text" name="Quantity" maxlength="5" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" onKeypress="AllowNumericOnly();">
			<input type="hidden" name="Accessory" value="<%=(rsAttachment.Fields.Item("insAttachment_id").Value)%>">		
		</td>
    </tr>
<%
			count++;
		}
	}
	rsAttachment.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="submit" value="Save" tabindex="10" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ArraySize" value="<%=count%>">
<input type="hidden" name="ClassID" value="<%=Request.QueryString("classid")%>">
<input type="hidden" name="MM_edit" value="true">
</form>
</body>
</html>
<%
rsAttachment.Close();
%>