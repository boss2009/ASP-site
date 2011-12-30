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
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");		
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");			
	var JobTitle = String(Request.Form("JobTitle")).replace(/'/g, "''");			
	var IsClerk = ((Request.Form("IsClerk")=="1") ? "1":"0");
	var IsCoordinator = ((Request.Form("IsCoordinator")=="1") ? "1":"0");
	var IsManager = ((Request.Form("IsManager")=="1") ? "1":"0");
	var IsRegionAdministrator = ((Request.Form("IsRegionAdministrator")=="1") ? "1":"0");
	var IsSystemAdministrator = ((Request.Form("IsSystemAdministrator")=="1") ? "1":"0");
	var IsConsultant = ((Request.Form("IsConsultant")=="1") ? "1":"0");
	var IsTechnician = ((Request.Form("IsTechnician")=="1") ? "1":"0");
	var IsSystemSupport = ((Request.Form("IsSystemSupport")=="1") ? "1":"0");							
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");				
	
	var cmdStaff = Server.CreateObject("ADODB.Command");
	cmdStaff.ActiveConnection = MM_cnnASP02_STRING;
	cmdStaff.CommandText = "dbo.cp_Staff2";
	cmdStaff.CommandType = 4;
	cmdStaff.CommandTimeout = 0;
	cmdStaff.Prepared = true;
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("RETURN_VALUE", 3, 4));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@intRecId", 3, 1,1,0));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@inspTitle_Typ_id", 2, 1,1,Request.Form("Title")));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvpFst_Name", 200, 1,50,FirstName));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvpLst_Name", 200, 1,50,LastName));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@inspRegion_Num", 2, 1,1,Request.Form("Region")));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvpNotes", 200, 1,256,Notes));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvpJobTitle", 200, 1,50,JobTitle));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Clerk", 2, 1,1,IsClerk));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Consultant", 2, 1,1,IsConsultant));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Coordinator", 2, 1,1,IsCoordinator));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Manager", 2, 1,1,IsManager));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_System_Support", 2, 1,1,IsSystemSupport));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Technican", 2, 1,1,IsTechnician));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Reg_Admin", 2, 1,1,IsRegionAdministrator));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@bitpIs_Sys_Admin", 2, 1,1,IsSystemAdministrator));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@inspSrtBy", 2, 1,1,1));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvFilter", 200, 1,1,""));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@insMode", 16, 1,1,0));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdStaff.Parameters.Append(cmdStaff.CreateParameter("@intRtnFlag", 3, 2));
	cmdStaff.Execute();	
	Response.Redirect("m002FS3.asp?insStaff_id="+cmdStaff.Parameters.Item("@intRtnFlag").Value);	
}

var rsRegion = Server.CreateObject("ADODB.Recordset");
rsRegion.ActiveConnection = MM_cnnASP02_STRING;
rsRegion.Source = "{call dbo.cp_Region}";
rsRegion.CursorType = 0;
rsRegion.CursorLocation = 2;
rsRegion.LockType = 3;
rsRegion.Open();

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
	<title>New Staff</title>
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
		if (!CheckTextArea(document.frm0101.Notes, 256)){
			alert("Text area cannot exceed 256 characters.");
			return ;
		}
	
		if (Trim(document.frm0101.LastName.value)==""){
			alert("Enter last name of staff.");
			document.frm0101.LastName.focus();
			return ;
		}
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="document.frm0101.Title.focus();">
<form action="<%=MM_insertAction%>" method="POST" name="frm0101">
<h5>New Staff</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Title:</td>
		<td nowrap><select name="Title" tabindex="1" accesskey="F">
			<% 
			while (!rsTitle.EOF) {
			%>
				<option value="<%=(rsTitle.Fields.Item("insTitle_Typ_id").Value)%>"><%=(rsTitle.Fields.Item("chvtitle").Value)%> 
			<% 
				rsTitle.MoveNext();
			} 
			%>
		</select></td>
		<td colspan="2" class="headrow">Job Role(s):</td>
	</tr>
	<tr> 
		<td nowrap>First Name:</td>
		<td nowrap><input type="text" name="FirstName" maxlength="50" value="" size="20" tabindex="2"></td>
		<td nowrap><input type="checkbox" name="IsClerk" value="1" tabindex="6" class="chkstyle">Clerk</td>
		<td nowrap><input type="checkbox" name="IsSystemSupport" value="1" tabindex="10" class="chkstyle">System Support</td>
    </tr>
    <tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" maxlength="50" value="" size="20" tabindex="3"></td>
		<td nowrap><input type="checkbox" name="IsConsultant" value="1" tabindex="7" class="chkstyle">Consultant</td>
		<td nowrap><input type="checkbox" name="IsTechnician" value="1" tabindex="11" class="chkstyle">Technician</td>
    </tr>
    <tr> 
		<td nowrap>Job Title:</td>
		<td nowrap><input type="text" name="JobTitle" maxlength="50" value="" size="20" tabindex="4"></td>
		<td nowrap><input type="checkbox" name="IsCoordinator" value="1" tabindex="8" class="chkstyle">Coordinator</td>
		<td nowrap><input type="checkbox" name="IsRegionAdministrator" value="1" tabindex="12" class="chkstyle">Region Administrator</td>
    </tr>
    <tr> 
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="5">
		<% 
		while (!rsRegion.EOF) { 			
		%>
			<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>"><%=(rsRegion.Fields.Item("chvname").Value)%> 
		<% 
			rsRegion.MoveNext();
		} 
		%>
		</select></td>
		<td nowrap><input type="checkbox" name="IsManager" value="1" tabindex="9" class="chkstyle">Manager</td>
		<td nowrap><input type="checkbox" name="IsSystemAdministrator" value="1" tabindex="13" class="chkstyle">System Administrator</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top" colspan="3"><textarea name="Notes" cols="65" rows="5" tabindex="14" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="15" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="16" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsRegion.Close();
rsTitle.Close();
%>