<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_update")) == "true") {	
	var Relationship = String(Request.Form("Relationship")).replace(/'/g, "''");	
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsRelationship = Server.CreateObject("ADODB.Recordset");
	rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
	rsRelationship.Source = "{call dbo.cp_Relationship2("+ Request.Form("MM_recordId") + ",'" + Request.Form("Description") + "'," + Request.Form("RelationshipType") + "," + IsActive + ",0,'E',0)}";
	rsRelationship.CursorType = 0;
	rsRelationship.CursorLocation = 2;
	rsRelationship.LockType = 3;
	rsRelationship.Open();
	Response.Redirect("m018q0314.asp");
}

var rsRelationship = Server.CreateObject("ADODB.Recordset");
rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
rsRelationship.Source = "{call dbo.cp_relationship2("+ Request.QueryString("insRtnship_id") + ",'',0,0,1,'Q',0)}";
rsRelationship.CursorType = 0;
rsRelationship.CursorLocation = 2;
rsRelationship.LockType = 3;
rsRelationship.Open();

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
	<title>Update Relationship Lookup</title>
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
<h5>Update Relationship Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsRelationship.Fields.Item("chvRtnship").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
    </tr>
	<tr>
		<td>Relationship Type:</td>		
		<td><select name="RelationshipType" tabindex="2">
			<%
			while (!rsRelationshipContactType.EOF){
			%>
				<option value="<%=rsRelationshipContactType.Fields.Item("insRelationship_Contact_Type_Id").Value%>" <%=((rsRelationship.Fields.Item("intobject_type_id").Value==rsRelationshipContactType.Fields.Item("insRelationship_Contact_Type_Id").Value)?"SELECTED":"")%>><%=rsRelationshipContactType.Fields.Item("chvRelationship_Contact_Type").Value%>
			<%
				rsRelationshipContactType.MoveNext();
			}
			%>
		</select></td>
	</tr>	
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" <%=((rsRelationship.Fields.Item("bitis_active").Value == 1)?"CHECKED":"")%> value="1" tabindex="3" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsRelationship.Fields.Item("insRtnship_id").Value %>">
</form>
</body>
</html>
<%
rsRelationship.Close();
rsRelationshipContactType.Close();
%>