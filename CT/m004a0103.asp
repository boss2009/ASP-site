<%@language="JAVASCRIPT"%> 
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");		
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");			
	var JobTitle = String(Request.Form("JobTitle")).replace(/'/g, "''");			
	var cmdInsertContact = Server.CreateObject("ADODB.Command");
	cmdInsertContact.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertContact.CommandText = "dbo.cp_Contacts";
	cmdInsertContact.CommandType = 4;
	cmdInsertContact.CommandTimeout = 0;
	cmdInsertContact.Prepared = true;
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@inspTitle_Typ_id", 2, 1,1,Request.Form("Title")));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@chvpFst_Name", 200, 1,20,FirstName));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@chvpLst_Name", 200, 1,20,LastName));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@chvJob_Title", 200, 1,50,JobTitle));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@bitIs_CmpyKey", 2, 1,1,0));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@insWork_type_id", 2, 1,1,Request.Form("WorkType")));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@insWork_id", 2, 1,1,Request.Form("WorkLocation")));	
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@inspSrtBy", 2, 1,1,0));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@inspSrtOrd", 2, 1,1,1));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@chvFilter", 200, 1,200,""));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertContact.Parameters.Append(cmdInsertContact.CreateParameter("@intRtnFlag", 3, 2));
	cmdInsertContact.Execute();

	var ContactID = cmdInsertContact.Parameters.Item("@intRtnFlag").Value;
	
	var rsUpdateEntryDate = Server.CreateObject("ADODB.Recordset");
	rsUpdateEntryDate.ActiveConnection = MM_cnnASP02_STRING;
	rsUpdateEntryDate.Source = "update tbl_contact set dtsEntrydate = '" + CurrentDate() + "' where intContact_id = " + ContactID;
	rsUpdateEntryDate.CursorType = 0;
	rsUpdateEntryDate.CursorLocation = 2;
	rsUpdateEntryDate.LockType = 3;
	rsUpdateEntryDate.Open();	
	
	switch (String(Request.Form("LinkToClass"))){
		//no link
		case "0":
		break;
		//client
		case "1":
			var rsLinkContact = Server.CreateObject("ADODB.Recordset");
			rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
			rsLinkContact.Source="{call dbo.cp_clnctact2("+Request.Form("LinkToObject")+","+ContactID+","+Request.Form("Relationship")+",0,0,'A',0)}";
			rsLinkContact.CursorType = 0;
			rsLinkContact.CursorLocation = 2;
			rsLinkContact.LockType = 3;
			rsLinkContact.Open();
		break;
		//company
		case "2":
			var rsLinkContact = Server.CreateObject("ADODB.Recordset");
			rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
			rsLinkContact.Source="{call dbo.cp_company_contact("+Request.Form("LinkToObject")+","+Request.Form("WorkType")+","+ContactID+",'A',0)}";
			rsLinkContact.CursorType = 0;
			rsLinkContact.CursorLocation = 2;
			rsLinkContact.LockType = 3;
			rsLinkContact.Open();
		break;	
		//institution
		case "3":
			var rsLinkContact = Server.CreateObject("ADODB.Recordset");
			rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
			rsLinkContact.Source="{call dbo.cp_school_Contacts("+Request.Form("LinkToObject")+","+ContactID+"," +Request.Form("Relationship")+",1,'A',0)}";
			rsLinkContact.CursorType = 0;
			rsLinkContact.CursorLocation = 2;
			rsLinkContact.LockType = 3;
			rsLinkContact.Open();
			break;			
		//on-site support
		case "3":
			var rsLinkContact = Server.CreateObject("ADODB.Recordset");
			rsLinkContact.ActiveConnection = MM_cnnASP02_STRING;
			rsLinkContact.Source="{call dbo.cp_pilat_site_support("+Request.Form("LinkToObject")+","+ContactID+",1,'A',0)}";
			rsLinkContact.CursorType = 0;
			rsLinkContact.CursorLocation = 2;
			rsLinkContact.LockType = 3;
			rsLinkContact.Open();
		break;			
	}	
	Response.Redirect("m004FS3.asp?intContact_id="+ContactID);
}

var rsWorkType = Server.CreateObject("ADODB.Recordset");
rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
rsWorkType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
rsWorkType.CursorType = 0;
rsWorkType.CursorLocation = 2;
rsWorkType.LockType = 3;
rsWorkType.Open();

var WorkType = ((String(Request.Form("WorkType"))=="")?13:Request.Form("WorkType"));

switch (String(WorkType)) {
	case "12":
		var rsWorkLocation = Server.CreateObject("ADODB.Recordset");
		rsWorkLocation.ActiveConnection = MM_cnnASP02_STRING;
		rsWorkLocation.Source = "{call dbo.cp_School}";
		rsWorkLocation.CursorType = 0;
		rsWorkLocation.CursorLocation = 2;
		rsWorkLocation.LockType = 3;
		rsWorkLocation.Open();
	break;
	default :
		var rsWorkLocation = Server.CreateObject("ADODB.Recordset");
		rsWorkLocation.ActiveConnection = MM_cnnASP02_STRING;
		rsWorkLocation.Source = "{call dbo.cp_get_company_work_type("+WorkType+",0)}";
		rsWorkLocation.CursorType = 0;
		rsWorkLocation.CursorLocation = 2;
		rsWorkLocation.LockType = 3;
		rsWorkLocation.Open();
	break;
}

var rsTitle = Server.CreateObject("ADODB.Recordset");
rsTitle.ActiveConnection = MM_cnnASP02_STRING;
rsTitle.Source = "{call dbo.cp_TITLE_type(0,0)}";
rsTitle.CursorType = 0;
rsTitle.CursorLocation = 2;
rsTitle.LockType = 3;
rsTitle.Open();

var rsRelationship = Server.CreateObject("ADODB.Recordset");
rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
rsRelationship.Source = "{call dbo.cp_Relationship(0,'',1,0,'Q',0)}";
rsRelationship.CursorType = 0;
rsRelationship.CursorLocation = 2;
rsRelationship.LockType = 3;
rsRelationship.Open();
%>
<html>
<head>
	<title>New Contact</title>
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
		if (document.frm0103.WorkLocation.value <= 0) {
			alert("Select a Work Location.");
			document.frm0103.WorkLocation.focus();
			return ;
		}
		if (Trim(document.frm0103.LastName.value)==""){
			alert("Enter Last Name.");
			document.frm0103.LastName.focus();
			return ;
		}
		document.frm0103.MM_insert.value = "true";
		document.frm0103.submit();
	}
	</script>
</head>
<body onLoad="document.frm0103.Title.focus();">
<form action="<%=MM_updateAction%>" method="POST" name="frm0103">
<h5>New Contact</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Title:</td>
		<td nowrap><select name="Title" tabindex="1" accesskey="F">
		<% 
		while (!rsTitle.EOF) { 			
		%>
			<option value="<%=(rsTitle.Fields.Item("insTitle_Typ_id").Value)%>" <%=((rsTitle.Fields.Item("insTitle_Typ_id").Value == Request.Form("Title"))?"SELECTED":"")%>><%=(rsTitle.Fields.Item("chvtitle").Value)%> 
		<% 
			rsTitle.MoveNext();
		} 
		%>
        </select></td>
	</tr>
    <tr> 
		<td nowrap>First Name:</td>
		<td nowrap><input type="text" name="FirstName" maxlength="20" value="<%=Request.Form("FirstName")%>" size="20" tabindex="2"></td>
    </tr>
    <tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" maxlength="20" value="<%=Request.Form("LastName")%>" size="20" tabindex="3"></td>
    </tr>
    <tr> 
		<td nowrap>Job Title:</td>
		<td nowrap><input type="text" name="JobTitle" maxlength="50" value="<%=Request.Form("JobTitle")%>" size="30" tabindex="4"></td>
    </tr>
    <tr> 
		<td nowrap>Work Type:</td>
		<td nowrap><select name="WorkType" tabindex="5" onChange="document.frm0103.submit();">
		<% 
		while (!rsWorkType.EOF) { 			
		%>
			<option value="<%=(rsWorkType.Fields.Item("intWork_type_id").Value)%>" <%=((rsWorkType.Fields.Item("intWork_type_id").Value == Request.Form("WorkType"))?"SELECTED":"")%>><%=(rsWorkType.Fields.Item("chvWork_type_desc").Value)%> 
		<% 
			rsWorkType.MoveNext();
		} 
		%>
        </select></td>		
    </tr>
    <tr> 
		<td nowrap>Work Location:</td>
		<td nowrap><select name="WorkLocation" tabindex="6">
		<% 
		switch (String(WorkType)) {
			case "12":
				while (!rsWorkLocation.EOF) {
		%>
					<option value="<%=(rsWorkLocation.Fields.Item("insSchool_id").Value)%>"><%=(rsWorkLocation.Fields.Item("chvName").Value)%></option>
		<%
					rsWorkLocation.MoveNext();
				}
			break;
			default :
				while (!rsWorkLocation.EOF) {
		%>
					<option value="<%=(rsWorkLocation.Fields.Item("intCompany_id").Value)%>"><%=(rsWorkLocation.Fields.Item("chvOrg_Name").Value)%></option>
		<%
					rsWorkLocation.MoveNext();
				}
			break;
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Relationship:</td>
		<td nowrap><select name="Relationship" tabindex="7" accesskey="L">
			<% 
			while (!rsRelationship.EOF) { 			
			%>
		        <option value="<%=(rsRelationship.Fields.Item("insRtnship_id").Value)%>" <%=((rsRelationship.Fields.Item("insRtnship_id").Value == Request.Form("Relationship"))?"SELECTED":"")%>><%=(rsRelationship.Fields.Item("chvname").Value)%> 
        	<% 
				rsRelationship.MoveNext();
			} 
			%>
        </select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="9" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="false">
<input type="hidden" name="LinkToClass" value="<%=Request.Form("LinkToClass")%>">
<input type="hidden" name="LinkToObject" value="<%=Request.Form("LinkToObject")%>">
</form>
</body>
</html>
<%
rsRelationship.Close();
rsWorkType.Close();
rsTitle.Close();
%>