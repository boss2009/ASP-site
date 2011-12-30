<%@language="JAVASCRIPT"%> 
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert"))=="true") {
	var rsInsertEquipment = Server.CreateObject("ADODB.Recordset");
	rsInsertEquipment.ActiveConnection = MM_cnnASP02_STRING;
	rsInsertEquipment.Source = "{call dbo.cp_company_equipment("+Request.Form("")+")}";
	rsInsertEquipment.CursorType = 0;
	rsInsertEquipment.CursorLocation = 2;
	rsInsertEquipment.LockType = 3;
	rsInsertEquipment.Open();
	Response.Redirect("");
}
%>
<html>
<head>
	<title>Add Equipment</title>
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
		if (document.frm06s01.ClassID.value==""){
			alert("Select Class.");
			document.frm06s01.List.focus();
			return ;
		}
		document.frm06s01.submit();
	}	
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=10,top=10,status=1");
		return ;
	}	   	
	</script>
</head>
<body onLoad="document.frm06s01.ClassName.focus();">
<form action="<%=MM_updateAction%>" method="POST" name="frm06s01">
<h5>Add Equipment</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Class Name:</td>
		<td nowrap><input type="text" name="ClassName" tabindex="1" accesskey="F" size="50" readonly></td>		
		<td nowrap><input type="button" name="List" value="List" onClick="openWindow('m006p01FS.asp','ClassSearch');" tabindex="2" accesskey="L" class="btnstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m006q0401.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>';" tabindex="4" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ClassID" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>