<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_Insert")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
	rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
	rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn2(0,"+ Request.Form("DocumentType") + ",'" + Description + "',0,'A',0)}";
	rsDocumentCondition.CursorType = 0;
	rsDocumentCondition.CursorLocation = 2;
	rsDocumentCondition.LockType = 3;
	rsDocumentCondition.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsDocumentType = Server.CreateObject("ADODB.Recordset");
rsDocumentType.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentType.Source = "{call dbo.cp_doc_type(0,'',0,0,0,0,0,'Q',0)}";
rsDocumentType.CursorType = 0;
rsDocumentType.CursorLocation = 2;
rsDocumentType.LockType = 3;
rsDocumentType.Open();
%>
<html>
<head>
	<title>New Document Condition Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0342.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0342.Description.value)==""){
			alert("Enter Description.");
			document.frm0342.Description.focus();
			return ;		
		}
		document.frm0342.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0342.Description.focus();">
<form name="frm0342" method="POST" action="<%=MM_editAction%>">
<h5>New Document Condition Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top">Description:</td>
		<td valign="top"><textarea name="Description" cols="65" rows="3" tabindex="1" accesskey="F"></textarea></td>
    </tr>
    <tr>
		<td>Document Type:</td> 
		<td><select name="DocumentType" tabindex="2" accesskey="L">		
		<%
		while (!rsDocumentType.EOF) {
		%>
			<option value="<%=rsDocumentType.Fields.Item("intDoc_Type_Id").Value%>"><%=rsDocumentType.Fields.Item("chvDoc_Type_Desc").Value%>
		<%
			rsDocumentType.MoveNext();
		}
		%>
		</select></td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>
<%
rsDocumentType.Close();
%>