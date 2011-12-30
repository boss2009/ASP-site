<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
}

var rsAttachment = Server.CreateObject("ADODB.Recordset");
rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
rsAttachment.Source = "{call dbo.cp_Get_Ivtry_Attachment("+ Request.QueryString("InventoryID") + ",0)}";
rsAttachment.CursorType = 0;
rsAttachment.CursorLocation = 2;
rsAttachment.LockType = 3;
rsAttachment.Open();
%>
<html>
<head>
	<title>Accessories</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
</head>
<body>
<table cellspacing="1" cellpadding="1" align="center">
	<tr> 
		<th class="headrow" nowrap align="left" width="260">Accessory</th>	
		<th class="headrow" nowrap align="left">Quantity</th>
    </tr>
<% 
var count = 0;
while (!rsAttachment.EOF) { 
	if (String(Request.QueryString("All"))=="1"){ 
%>
    <tr> 
		<td><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%></td>
		<td><input type="text" name="Quantity" maxlength="5" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" readonly></td>
    </tr>
<%	
		count++;
	} else {
		if (String(rsAttachment.Fields.Item("insQuantity").Value)!="0") { 
%>
    <tr> 
		<td><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%></td>
		<td><input type="text" name="Quantity" maxlength="5" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" readonly></td>
    </tr>
<%	
			count++;
		}
	}
	rsAttachment.MoveNext();
}
%>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
</body>
</html>
<%
rsAttachment.Close();
%>