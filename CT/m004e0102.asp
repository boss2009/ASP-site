<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true") {
	switch (String(Request.Form("LinkToClass"))){
		//client
		case "1":
			var rsChangeRelationship = Server.CreateObject("ADODB.Recordset");
			rsChangeRelationship.ActiveConnection = MM_cnnASP02_STRING;
			rsChangeRelationship.Source="{call dbo.cp_ctc_relationship2("+Request.QueryString("intContact_id")+",'2',"+Request.Form("LinkToObject")+","+Request.Form("NewRelationship")+",0,'E',0)}";
			rsChangeRelationship.CursorType = 0;
			rsChangeRelationship.CursorLocation = 2;
			rsChangeRelationship.LockType = 3;
			rsChangeRelationship.Open();
		break;
		//institution
		case "3":
			var rsChangeRelationship = Server.CreateObject("ADODB.Recordset");
			rsChangeRelationship.ActiveConnection = MM_cnnASP02_STRING;
			rsChangeRelationship.Source="{call dbo.cp_ctc_relationship2("+Request.QueryString("intContact_id")+",'1',"+Request.Form("LinkToObject")+","+Request.Form("NewRelationship")+",0,'E',0)}";
			rsChangeRelationship.CursorType = 0;
			rsChangeRelationship.CursorLocation = 2;
			rsChangeRelationship.LockType = 3;
			rsChangeRelationship.Open();
		break;
	}
	Response.Redirect("InsertSuccessful.html");
}

switch (String(Request.QueryString("LinkToClass"))){
	//client
	case "1":
		var ObjectType = "Client";
		var rsRelationship = Server.CreateObject("ADODB.Recordset");
		rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
		rsRelationship.Source = "{call dbo.cp_relationship2(0,'',3,1,0,'Q',0)}";
		rsRelationship.CursorType = 0;
		rsRelationship.CursorLocation = 2;
		rsRelationship.LockType = 3;
		rsRelationship.Open();		
	break;
	//institution
	case "3":
		var ObjectType = "Organization";	
		var rsRelationship = Server.CreateObject("ADODB.Recordset");
		rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
		rsRelationship.Source = "{call dbo.cp_relationship2(0,'',12,1,0,'Q',0)}";
		rsRelationship.CursorType = 0;
		rsRelationship.CursorLocation = 2;
		rsRelationship.LockType = 3;
		rsRelationship.Open();		
	break;				
}
%>
<html>
<head>
	<title>Change Contact Relationship</title>
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
			case 85:
				//alert("U");
				document.frm0102.reset();
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
		document.frm0102.submit();
	}
	</script>	
</head>
<body>
<form action="<%=MM_updateAction%>" method="POST" name="frm0102">
<h5>Change Contact Relationship</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Object Type:</td>
		<td nowrap><input type="text" name="ObjectType" value="<%=ObjectType%>" size="20" readonly tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Object Name:</td>
		<td nowrap><input type="text" name="ObjectName" value="<%=Request.QueryString("ObjectName")%>" size="30" readonly tabindex="2"></td>
	</tr>
	<tr>
		<td nowrap>Contact Name:</td>
		<td nowrap><input type="text" name="ContactName" value="<%=Request.QueryString("ContactName")%>" size="30" readonly tabindex="3"></td>
	</tr>
	<tr>
		<td nowrap>Current Relationship:</td>
		<td nowrap><input type="text" name="CurrentRelationship" value="<%=Request.QueryString("Relationship")%>" size="30" readonly tabindex="4"></td>
	</tr>
	<tr>
		<td nowrap>New Relationship:</td>
		<td nowrap><select name="NewRelationship" tabindex="5" accesskey="L">
			<%
			while (!rsRelationship.EOF){
			%>
				<option value="<%=rsRelationship.Fields.Item("insRtnship_id").Value%>"><%=rsRelationship.Fields.Item("chvRtnship").Value%>
			<%
				rsRelationship.MoveNext();
			}
			%>
				<option value="0">Not Available								
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="7" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="LinkToClass" value="<%=((String(Request.QueryString("LinkToClass"))=="undefined")?"0":Request.QueryString("LinkToClass"))%>">
<input type="hidden" name="LinkToObject" value="<%=Request.QueryString("LinkToObject")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsRelationship.Close();
%>