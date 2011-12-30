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
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsRelationship = Server.CreateObject("ADODB.Recordset");
	rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
	rsRelationship.Source = "{call dbo.cp_Relationship2(0,'"+ Description + "'," + Request.Form("RelationshipType") + "," + IsActive + ",0,'A',0)}";
	rsRelationship.CursorType = 0;
	rsRelationship.CursorLocation = 2;
	rsRelationship.LockType = 3;
	rsRelationship.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsRelationshipContactType = Server.CreateObject("ADODB.Recordset");
rsRelationshipContactType.ActiveConnection = MM_cnnASP02_STRING;
rsRelationshipContactType.Source = "{call dbo.cp_relationship_contact_type(0,'',0,0,'Q',0)}";
rsRelationshipContactType.CursorType = 0;
rsRelationshipContactType.CursorLocation = 2;
rsRelationshipContactType.LockType = 3;
rsRelationshipContactType.Open();
%>
<html>
<head>
	<title>New Relationship Lookup</title>
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
				document.frm0314.reset();
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
		if (Trim(document.frm0314.Description.value)==""){
			alert("Enter Description.");
			document.frm0314.Description.focus();
			return ;		
		}
		document.frm0314.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0314.Description.focus();">
<form name="frm0314" method="POST" action="<%=MM_editAction%>">
<h5>New Relationship Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
    </tr>
	<tr>
		<td nowrap>Relationship Type:</td>		
		<td nowrap><select name="RelationshipType" tabindex="2">
		<%
		while (!rsRelationshipContactType.EOF){
		%>
			<option value="<%=rsRelationshipContactType.Fields.Item("insRelationship_Contact_Type_Id").Value%>"><%=rsRelationshipContactType.Fields.Item("chvRelationship_Contact_Type").Value%>
		<%
			rsRelationshipContactType.MoveNext();
		}
		%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" value="1" tabindex=""3" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>
<%
rsRelationshipContactType.Close();
%>