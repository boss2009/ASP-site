<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (Request.Form("MM_update") == "true"){
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");		
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");			
	var JobTitle = String(Request.Form("JobTitle")).replace(/'/g, "''");			
	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_contacts(" + Request.Form("ContactID") + ","+Request.Form("Title")+",'" + FirstName + "','" + LastName + "','" + JobTitle +"',0,"+Request.Form("WorkType")+","+Request.Form("WorkPlace")+",0,0,'',0,'E',0)}";
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m004e0101.asp&intContact_id="+Request.Form("ContactID"));
}

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_contacts("+Request.QueryString("intContact_id")+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();	

var rsWorkType = Server.CreateObject("ADODB.Recordset");
rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
rsWorkType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
rsWorkType.CursorType = 0;
rsWorkType.CursorLocation = 2;
rsWorkType.LockType = 3;
rsWorkType.Open();

var WorkType = ((String(Request.Form("WorkType"))=="undefined")?rsContact.Fields.Item("intWork_type_id").Value:Request.Form("WorkType"));

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
%>
<html>
<head>
	<title>General Information</title>
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
				document.frm0101.reset();
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
		if (document.frm0101.WorkPlace.value <= 0) {
			alert("Select work place.");
			document.frm0101.WorkPlace.focus();
			return ;
		}
		if (Trim(document.frm0101.LastName.value)==""){
			alert("Enter Last Name.");
			document.frm0101.LastName.focus();
			return ;
		}
		document.frm0101.MM_update.value="true";
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="document.frm0101.Title.focus();"> 
<form action="<%=MM_updateAction%>" method="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Title:</td>
		<td nowrap><select name="Title" tabindex="1" accesskey="F">
		<% 
		while (!rsTitle.EOF) { 			
		%>
			<option value="<%=(rsTitle.Fields.Item("insTitle_Typ_id").Value)%>" <%=((rsContact.Fields.Item("insTitle_Typ_id").Value==rsTitle.Fields.Item("insTitle_Typ_id").Value)?"SELECTED":"")%>><%=(rsTitle.Fields.Item("chvtitle").Value)%> 
		<% 
			rsTitle.MoveNext();
		} 
		%>		
		</select></td>
	</tr>		
	<tr>
		<td nowrap>First Name:</td>
		<td nowrap><input type="text" name="FirstName" value="<%=rsContact.Fields.Item("chvFst_Name").Value%>" maxlength="50" size="20" tabindex="2"></td>
	</tr>
	<tr>
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" value="<%=rsContact.Fields.Item("chvLst_Name").Value%>" maxlength="50" size="20" tabindex="3"></td>
	</tr>
	<tr>
		<td nowrap>Job Title:</td>
		<td nowrap><input type="text" name="JobTitle" value="<%=(rsContact.Fields.Item("chvJob_Title").Value)%>" maxlength="50" size="30" tabindex="4"></td>
	</tr>
	<tr>
		<td nowrap>Work Type:</td>
		<td nowrap><select name="WorkType" tabindex="5" onChange="document.frm0101.submit();">
		<% 
		while (!rsWorkType.EOF) { 
		%>
			<option value="<%=(rsWorkType.Fields.Item("intWork_type_id").Value)%>" <%if (String(Request.Form("WorkType"))=="undefined") {Response.Write((rsContact.Fields.Item("intWork_type_id").Value==rsWorkType.Fields.Item("intWork_type_id").Value)?"SELECTED":"")} else {Response.Write((Request.Form("WorkType")==rsWorkType.Fields.Item("intWork_type_id").Value)?"SELECTED":"")}%>><%=(rsWorkType.Fields.Item("chvWork_type_desc").Value)%> 
		<%
			rsWorkType.MoveNext();
		}
		%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Work Place:</td>		
		<td nowrap><select name="WorkPlace" tabindex="6" accesskey="L">
		<% 
		switch (String(WorkType)) {
			case "12":
				while (!rsWorkLocation.EOF) {
		%>
					<option value="<%=(rsWorkLocation.Fields.Item("insSchool_id").Value)%>" <%if (String(Request.Form("WorkPlace"))=="undefined") {Response.Write((rsContact.Fields.Item("insWork_id").Value==rsWorkLocation.Fields.Item("insSchool_id").Value)?"SELECTED":"")} else {Response.Write((rsWorkLocation.Fields.Item("insSchool_id").Value == Request.Form("WorkPlace"))?"SELECTED":"")}%>><%=(rsWorkLocation.Fields.Item("chvName").Value)%></option>
		<%
					rsWorkLocation.MoveNext();
				}
			break;
			default :
				while (!rsWorkLocation.EOF) {
		%>
					<option value="<%=(rsWorkLocation.Fields.Item("intCompany_id").Value)%>" <%if (String(Request.Form("WorkPlace"))=="undefined") {Response.Write((rsContact.Fields.Item("insWork_id").Value==rsWorkLocation.Fields.Item("intCompany_id").Value)?"SELECTED":"")} else {Response.Write((rsWorkLocation.Fields.Item("intCompany_id").Value == Request.Form("WorkPlace"))?"SELECTED":"")}%>><%=(rsWorkLocation.Fields.Item("chvOrg_Name").Value)%></option>
		<%
					rsWorkLocation.MoveNext();
				}
			break;
		}
		%>
        </select></td>
    </tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="7" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="8" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ContactID" value="<%=Request.QueryString("intContact_id")%>">
<input type="hidden" name="MM_update" value="false">
</form>
</body>
</html>
<%
rsContact.Close();
rsWorkType.Close();
rsTitle.Close();
%>