<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var InstitutionName = String(Request.Form("InstitutionName")).replace(/'/g, "''");		
	var rsInstitution = Server.CreateObject("ADODB.Recordset");
	rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitution.Source = "{call dbo.cp_school2("+Request.Form("InstitutionID") + ",'" + InstitutionName + "'," + Request.Form("Region") + "," + Request.Form("Type") +","+Request.Form("ParentSchool")+"," + Request.Form("CampusType") + "," + Session("insStaff_id") + ",0,0,'',0,'E',0)}";
	rsInstitution.CursorType = 0;
	rsInstitution.CursorLocation = 2;
	rsInstitution.LockType = 3;
	rsInstitution.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012e0101.asp&insSchool_id="+Request.Form("InstitutionID"));
}

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2("+Request.QueryString("insSchool_id")+",'',0,0,0,0,0,0,0,'',1,'Q',0)}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();	

var rsRegion = Server.CreateObject("ADODB.Recordset");
rsRegion.ActiveConnection = MM_cnnASP02_STRING;
rsRegion.Source = "{call dbo.cp_asp_lkup(7)}";
rsRegion.CursorType = 0;
rsRegion.CursorLocation = 2;
rsRegion.LockType = 3;
rsRegion.Open();

var rsType = Server.CreateObject("ADODB.Recordset");
rsType.ActiveConnection = MM_cnnASP02_STRING;
rsType.Source = "{call dbo.cp_school_type(0,'',1,0,'Q',0)}";
rsType.CursorType = 0;
rsType.CursorLocation = 2;
rsType.LockType = 3;
rsType.Open();

var rsParentSchool = Server.CreateObject("ADODB.Recordset");
rsParentSchool.ActiveConnection = MM_cnnASP02_STRING;
rsParentSchool.Source = "{call dbo.cp_school2(0,'',0,0,0,0,0,1,0,'',2,'Q',0)}";
rsParentSchool.CursorType = 0;
rsParentSchool.CursorLocation = 2;
rsParentSchool.LockType = 3;
rsParentSchool.Open();
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
	function Init() {
		ChangeCampusType();
		document.frm0101.InstitutionName.focus();
	}
	
	function ChangeCampusType(){
		if (document.frm0101.CampusType.value=="1"){
			ParentSchoolLabel.style.visibility = "hidden";
			document.frm0101.ParentSchool.style.visibility = "hidden";		
		} else {
			ParentSchoolLabel.style.visibility = "visible";
			document.frm0101.ParentSchool.style.visibility = "visible";
		}
	}

	function Save(){
		if (Trim(document.frm0101.InstitutionName.value)==""){
			alert("Enter Institution Name.");
			document.frm0101.InstitutionName.focus();
			return ;
		}
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="Init();"> 
<form action="<%=MM_updateAction%>" method="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Institution Name:</td>
		<td nowrap><input type="text" name="InstitutionName" maxlength="50" value="<%=rsInstitution.Fields.Item("chvSchool_Name").Value%>" size="50" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="2">
            <% 
			while (!rsRegion.EOF) { 			
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsInstitution.Fields.Item("insRegion_Num").Value==rsRegion.Fields.Item("insRegion_num").Value)?"SELECTED":"")%>><%=(rsRegion.Fields.Item("chvname").Value)%> 
			<% 
				rsRegion.MoveNext();
			} 
			%>		
		</select></td>	
	</tr>
	<tr>	
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="3">
            <% 
			while (!rsType.EOF) { 			
			%>
				<option value="<%=(rsType.Fields.Item("insSchool_type_id").Value)%>" <%=((rsInstitution.Fields.Item("insSchool_type_id").Value==rsType.Fields.Item("insSchool_type_id").Value)?"SELECTED":"")%>><%=(rsType.Fields.Item("chvSchool_Type").Value)%> 
			<% 
				rsType.MoveNext();
			} 
			%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Campus Type:</td>
		<td nowrap><select name="CampusType" tabindex="4" accesskey="L" onChange="ChangeCampusType();">
			<option value="1" <%=((rsInstitution.Fields.Item("bitIs_MainCampus").value=="1")?"SELECTED":"")%>>Main Campus
			<option value="0" <%=((rsInstitution.Fields.Item("bitIs_MainCampus").value=="0")?"SELECTED":"")%>>Satellite Campus
		</select></td>		
	</tr>
	<tr>
		<td nowrap><span id="ParentSchoolLabel">Parent School:</span></td>
		<td nowrap><select name="ParentSchool" tabindex="5">
			<% 
			while (!rsParentSchool.EOF) { 
			%>
		       <option value="<%=(rsParentSchool.Fields.Item("insSchool_id").Value)%>" <%=((rsParentSchool.Fields.Item("insSchool_id").Value==rsInstitution.Fields.Item("insSuper_School_id").Value)?"SELECTED":"")%>><%=(rsParentSchool.Fields.Item("chvSchool_Name").Value)%>
			<%
				rsParentSchool.MoveNext();
			}
			%>		
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="6" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="7" class="btnstyle"></td>		
		<td><input type="button" value="Close" tabindex="8" onClick="top.window.close();" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="InstitutionID" value="<%=Request.QueryString("insSchool_id")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsRegion.Close();
rsType.Close();
rsParentSchool.Close();
%>