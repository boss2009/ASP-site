<%@language="JAVASCRIPT"%> 
<!--#INCLUDE File="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_insertAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_insertAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var InstitutionName = String(Request.Form("InstitutionName")).replace(/'/g, "'");		
	var cmdInstitution = Server.CreateObject("ADODB.Command");
	cmdInstitution.ActiveConnection = MM_cnnASP02_STRING;
	cmdInstitution.CommandText = "dbo.cp_school2";
	cmdInstitution.CommandType = 4;
	cmdInstitution.CommandTimeout = 0;
	cmdInstitution.Prepared = true;
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@chvName", 200, 1,50,InstitutionName));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@insRegion_Num", 2, 1,1,Request.Form("Region")));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@insSchool_type_id", 2, 1,1,Request.Form("Type")));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@insSuper_School_id", 2, 1,1,Request.Form("ParentSchool")));			
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@bitIs_MainCampus", 2, 1,1,Request.Form("CampusType")));		
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@inspSrtBy", 2, 1,1,1));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@chvFilter", 200, 1,1,""));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@insMode", 16, 1,1,0));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInstitution.Parameters.Append(cmdInstitution.CreateParameter("@intRtnFlag", 3, 2));
	cmdInstitution.Execute();	
	Response.Redirect("m012FS3.asp?insSchool_id="+cmdInstitution.Parameters.Item("@intRtnFlag").Value);	
}

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
	<title>New Institution</title>
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
	<script language="JavaScript">
	function ChangeCampusType(){
		if (document.frm0101.CampusType.value==1){
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
<body onLoad="document.frm0101.InstitutionName.focus();">
<form action="<%=MM_insertAction%>" method="POST" name="frm0101">
<h5>New Institution</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Institution Name:</td>
		<td><input type="text" name="InstitutionName" maxlength="50" value="" size="50" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Region:</td>
		<td><select name="Region" tabindex="2">
            <% 
			while (!rsRegion.EOF) { 			
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>"><%=(rsRegion.Fields.Item("chvname").Value)%> 
			<% 
				rsRegion.MoveNext();
			} 
			%>		
		</select></td>	
	</tr>
	<tr>	
		<td nowrap>Type:</td>
		<td><select name="Type" tabindex="3">
            <% 
			while (!rsType.EOF) { 			
			%>
				<option value="<%=(rsType.Fields.Item("insSchool_type_id").Value)%>"><%=(rsType.Fields.Item("chvSchool_Type").Value)%> 
			<% 
				rsType.MoveNext();
			} 
			%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Campus Type:</td>
		<td nowrap><select name="CampusType" tabindex="4" accesskey="L" onChange="ChangeCampusType();" style="width: 200px">
			<option value="1">Main Campus
			<option value="0">Satellite Campus
		</select></td>
	</tr>
	<tr>
		<td nowrap><span id="ParentSchoolLabel" style="visibility: hidden">Parent School:</span></td>
		<td nowrap><select name="ParentSchool" tabindex="5" style="visibility: hidden; width: 200px">
		<% 
		while (!rsParentSchool.EOF) { 
		%>
		   <option value="<%=(rsParentSchool.Fields.Item("insSchool_id").Value)%>"><%=(rsParentSchool.Fields.Item("chvSchool_Name").Value)%>
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
		<td><input type="button" value="Save" onClick="Save();" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="6" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsRegion.Close();
rsType.Close();
rsParentSchool.Close();
%>