<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
}

if (Request.QueryString("MM_update") == "true"){
	var rsAttachment = Server.CreateObject("ADODB.Recordset");
	rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
	rsAttachment.CursorType = 0;
	rsAttachment.CursorLocation = 2;
	rsAttachment.LockType = 3;	
	for (var i=1; i<=Request.QueryString("ArraySize"); i++){
		rsAttachment.Source = "{call dbo.cp_Update_Ivtry_Attachment(" + Request.QueryString("intEquip_Set_id") + "," + Request.QueryString("Accessory")(i) + "," + Request.QueryString("Quantity")(i) + "," + Session("insStaff_id") + ",0)}";
		rsAttachment.Open();
		rsAttachment.Close();			
	}	
	Response.Redirect("UpdateSuccessful2.asp?page=m003e0201.asp&intEquip_Set_id="+Request.QueryString("intEquip_Set_id"));	
}

var rsAttachment = Server.CreateObject("ADODB.Recordset");
rsAttachment.ActiveConnection = MM_cnnASP02_STRING;
rsAttachment.Source = "{call dbo.cp_Get_Ivtry_Attachment("+ Request.QueryString("intEquip_Set_id") + ",0)}";
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
			case 83 :
				//alert("S");
				Save();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		document.frm0201.submit();
	}
	</script>
</head>
<body>
<form action="<%=MM_editAction%>" method="GET" name="frm0201">
<h5>Accessories</h5>
<a href="m003e0201.asp?All=1&intEquip_Set_id=<%=Request.QueryString("intEquip_Set_id")%>">Show All</a> | <a href="m003e0201.asp?All=0&intEquip_Set_id=<%=Request.QueryString("intEquip_Set_id")%>">Show Only Checked</a>
<hr>
<table cellspacing="1" cellpadding="1">
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
		<td nowrap><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%><input type="hidden" name="Accessory" value="<%=(rsAttachment.Fields.Item("insAttachment_id").Value)%>"></td>
		<td nowrap align="center"><input type="text" name="Quantity" maxlength="5" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" onKeypress="AllowNumericOnly();" ></td>
    </tr>
<%
		count++;
	} else {
		if (String(rsAttachment.Fields.Item("insQuantity").Value)!="0") { 
%>
	<tr> 
		<td nowrap><%=(rsAttachment.Fields.Item("chvAttach_Name").Value)%><input type="hidden" name="Accessory" value="<%=(rsAttachment.Fields.Item("insAttachment_id").Value)%>"></td>
		<td nowrap align="center"><input type="text" name="Quantity" maxlength="5" size="3" value="<%=(rsAttachment.Fields.Item("insQuantity").Value)%>" onKeypress="AllowNumericOnly();" ></td>
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
		<td><input type="button" value="Save" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="top.window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ArraySize" value="<%=count%>">
<input type="hidden" name="intEquip_Set_id" value="<%=Request.QueryString("intEquip_Set_id")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsAttachment.Close();
%>