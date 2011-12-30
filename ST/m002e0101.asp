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
	var rsStaff = Server.CreateObject("ADODB.Recordset");
	rsStaff.ActiveConnection = MM_cnnASP02_STRING;
	rsStaff.Source = "{call dbo.cp_staff2(" + Request.Form("StaffID") + "," + Request.Form("Title") + ",'" + FirstName + "','" + LastName + "'," + Request.Form("Region") + ",'" + Notes +"','" + JobTitle + "'," + Session("insStaff_id") + "," + IsClerk + "," + IsConsultant + "," + IsCoordinator + "," + IsManager + "," + IsSystemSupport + "," + IsTechnician + "," + IsRegionAdministrator + "," + IsSystemAdministrator + ",0,0,'',0,'E',0)}";
	rsStaff.CursorType = 0;
	rsStaff.CursorLocation = 2;
	rsStaff.LockType = 3;
	rsStaff.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m002e0101.asp&insStaff_id="+Request.Form("StaffID"));
}

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_staff2("+Request.QueryString("insStaff_id")+",0,'','',0,'','',0,0,0,0,0,0,0,0,0,1,0,'',1,'Q',0)}"
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();	

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
		if (!CheckTextArea(document.frm0101.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (Trim(document.frm0101.LastName.value)==""){
			alert("Enter Staff's Last Name.");
			document.frm0101.LastName.focus();
			return ;
		}
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
				<option value="<%=(rsTitle.Fields.Item("insTitle_Typ_id").Value)%>" <%=((rsStaff.Fields.Item("insTitle_Typ_id").Value==rsTitle.Fields.Item("insTitle_Typ_id").Value)?"SELECTED":"")%>><%=(rsTitle.Fields.Item("chvtitle").Value)%> 
			<% 
				rsTitle.MoveNext();
			} 
			%>		
		</select></td>
		<td colspan="2" class="headrow">Job Role(s):</td>			
	<tr>
		<td nowrap>First Name:</td>
		<td nowrap><input type="text" name="FirstName" maxlength="50" value="<%=rsStaff.Fields.Item("chvFst_Name").Value%>" size="20" tabindex="2"></td>
		<td nowrap><input type="checkbox" name="IsClerk" <%=((rsStaff.Fields.Item("bitIs_Clerk").value=="1")?"CHECKED":"")%> value="1" tabindex="6" class="chkstyle">Clerk</td>
		<td nowrap><input type="checkbox" name="IsSystemSupport" <%=((rsStaff.Fields.Item("bitIs_System_Support").value=="1")?"CHECKED":"")%> value="1" tabindex="10" class="chkstyle">System Support</td>						
	</tr>
	<tr>
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" maxlength="50" value="<%=rsStaff.Fields.Item("chvLst_Name").Value%>" size="20" tabindex="3"></td>
		<td nowrap><input type="checkbox" name="IsConsultant" <%=((rsStaff.Fields.Item("bitIs_Consultant").value=="1")?"CHECKED":"")%> value="1" tabindex="7" class="chkstyle">Consultant</td>
		<td nowrap><input type="checkbox" name="IsTechnician" <%=((rsStaff.Fields.Item("bitIs_Technican").value=="1")?"CHECKED":"")%> value="1" tabindex="11" class="chkstyle">Technician</td>						
	</tr>
	<tr>
		<td nowrap>Job Title:</td>
		<td nowrap><input type="text" name="JobTitle" maxlength="50" value="<%=(rsStaff.Fields.Item("chvJobTitle").Value)%>" size="30" tabindex="4"></td>
		<td nowrap><input type="checkbox" name="IsCoordinator" <%=((rsStaff.Fields.Item("bitIs_Coordinator").value=="1")?"CHECKED":"")%> value="1" tabindex="8" class="chkstyle">Coordinator</td>
		<td nowrap><input type="checkbox" name="IsRegionAdministrator" <%=((rsStaff.Fields.Item("bitIs_Reg_Admin").value=="1")?"CHECKED":"")%> value="1" tabindex="12" class="chkstyle">Region Administrator</td>						
	</tr>	
	<tr>
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="5">
            <% 
			while (!rsRegion.EOF) { 			
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsStaff.Fields.Item("insRegion_Num").Value==rsRegion.Fields.Item("insRegion_num").Value)?"SELECTED":"")%>><%=(rsRegion.Fields.Item("chvname").Value)%> 
			<% 
				rsRegion.MoveNext();
			} 
			%>		
		</select></td>	
		<td nowrap><input type="checkbox" name="IsManager" <%=((rsStaff.Fields.Item("bitIs_Manager").value=="1")?"CHECKED":"")%> value="1" tabindex="9" class="chkstyle">Manager</td>	
		<td nowrap><input type="checkbox" name="IsSystemAdministrator" <%=((rsStaff.Fields.Item("bitIs_Sys_Admin").value=="1")?"CHECKED":"")%> value="1" tabindex="13" class="chkstyle">System Administrator</td>		
	</tr>
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top" colspan="3"><textarea name="Notes" cols="65" rows="5" tabindex="14" accesskey="L"><%=rsStaff.Fields.Item("chvStaff_Notes").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="15" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="17" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="StaffID" value="<%=Request.QueryString("insStaff_id")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsStaff.Close();
rsRegion.Close();
rsTitle.Close();
%>