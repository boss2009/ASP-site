<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");	
	var MiddleName = String(Request.Form("MiddleName")).replace(/'/g, "''");	
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");
	var SetBCServed = ((Request.Form("SetBCServed")=="1")?"1":"0");	
	var PRCVIServed = ((Request.Form("PRCVIServed")=="1")?"1":"0");			
	var IsFirstNations = ((Request.Form("IsFirstNations")=="1")?"1":"0");				
	var cmdInsertClient = Server.CreateObject("ADODB.Command");
	cmdInsertClient.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertClient.CommandText = "dbo.cp_Adult_Client5A";
	cmdInsertClient.CommandType = 4;
	cmdInsertClient.CommandTimeout = 0;
	cmdInsertClient.Prepared = true;
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@intpID", 3, 1,1,0));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chvFst_name", 200, 1,20,FirstName));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chvMdl_name", 200, 1,20,MiddleName));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chvLst_name", 200, 1,20,LastName));		
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chrSIN_no", 129, 1,20,Request.Form("SIN")));			
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chrPEN_no", 129, 1,20,Request.Form("PEN")));				
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@bitGender_is_male", 2, 1,1,Request.Form("Gender")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@dtsBirth_date", 200, 1,30,Request.Form("DateOfBirth")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insRegion_num", 2, 1,1,Request.Form("Region")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insStdnt_Status_id", 2, 1,1,Request.Form("Status")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insCase_mngr_id", 2, 1,1,Request.Form("CaseManager")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insDsbty1_id", 2, 1,1,Request.Form("PrimaryDisability")));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insDsbty2_id", 2, 1,1,Request.Form("SecondaryDisability")));	
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@bitIsDefault_asp", 2, 1,1,Request.Form("ProgramStanding")));	
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@bitIs_Prx_SETBC", 2, 1,1,SetBCServed));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@bitIS_Prx_PRCVI", 2, 1,1,PRCVIServed));		
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@inspSrtBy", 2, 1,1,1));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chvFilter", 200, 1,1,""));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@bitIs_Prx_SETBC", 2, 1,1,IsFirstNations));	
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@chvTask", 129, 1,1,"A"));
	cmdInsertClient.Parameters.Append(cmdInsertClient.CreateParameter("@intRtnValue", 2, 2));
	cmdInsertClient.Execute();
	Response.Redirect("m001FS3.asp?intAdult_id="+cmdInsertClient.Parameters.Item("@intRtnValue").Value);
}

var rsRegion = Server.CreateObject("ADODB.Recordset");
rsRegion.ActiveConnection = MM_cnnASP02_STRING;
rsRegion.Source = "{call dbo.cp_Region}";
rsRegion.CursorType = 0;
rsRegion.CursorLocation = 2;
rsRegion.LockType = 3;
rsRegion.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_ASP_Lkup2(8,0,'',0,'1', 0)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();

var rsDisability = Server.CreateObject("ADODB.Recordset");
rsDisability.ActiveConnection = MM_cnnASP02_STRING;
rsDisability.Source = "{call dbo.cp_ASP_Lkup2(9, 0, '', 0, '1',0)}";
rsDisability.CursorType = 0;
rsDisability.CursorLocation = 2;
rsDisability.LockType = 3;
rsDisability.Open();

var rsCaseManager = Server.CreateObject("ADODB.Recordset");
rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
rsCaseManager.CursorType = 0;
rsCaseManager.CursorLocation = 2;
rsCaseManager.LockType = 3;
rsCaseManager.Open();
%>
<html>
<head>
	<title>New Client</title>
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
		if (!CheckSIN(document.frm0101.Sin.value)) {
			alert("Invalid Social Insurance Number.");
			document.frm0101.Sin.focus();
			return ;
		}
		if (!CheckDate(document.frm0101.DateOfBirth.value)){
			alert("Invalid Date of Birth.");
			document.frm0101.DateOfBirth.focus();
			return ;
		}
		if (Trim(document.frm0101.FirstName.value)=="") {
			alert("Enter First Name.");
			document.frm0101.FirstName.focus();
			return ;
		}
		if (Trim(document.frm0101.LastName.value)=="") {
			alert("Enter Last Name.");
			document.frm0101.LastName.focus();
			return ;
		}
		document.frm0101.Sin.value = LeaveDigits(document.frm0101.Sin.value);
		document.frm0101.submit();
		document.frm0101.btnSave.disabled = true;
	}
	</script>
</head>
<body onLoad="javascript:document.frm0101.FirstName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0101">
<h5>New Client:</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>First Name:</td>
		<td nowrap width="200"><input type="text" name="FirstName" maxlength="20" tabindex="1" accesskey="F"></td>
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="9">
			<% 
			while (!rsRegion.EOF) {
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsRegion.Fields.Item("insRegion_num").Value == 0)?"SELECTED":"")%>><%=(rsRegion.Fields.Item("chvname").Value)%></option>
			<%
				rsRegion.MoveNext();
			}
			%>
		</select></td>		
    </tr>
    <tr> 
		<td nowrap>Middle Name:</td>
		<td nowrap><input type="text" name="MiddleName" maxlength="20" tabindex="2"></td>
		<td nowrap>Status:</td>
		<td nowrap><select name="Status" tabindex="10" style="width: 150px">
			<% 
			while (!rsStatus.EOF) {
			%>
				<option value="<%=(rsStatus.Fields.Item("insStdnt_status_id").Value)%>" <%=((rsStatus.Fields.Item("insStdnt_status_id").Value == 10)?"SELECTED":"")%>><%=(rsStatus.Fields.Item("chvName").Value)%></option>
			<%
				rsStatus.MoveNext();
			}
			%>
		</select></td>		
    </tr>
    <tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" maxlength="20" tabindex="3"></td>
		<td nowrap>Case Manager:</td>
		<td nowrap><select name="CaseManager" tabindex="11" style="width: 150px">
			<% 
			while (!rsCaseManager.EOF) {
			%>
				<option value="<%=(rsCaseManager.Fields.Item("insId").Value)%>" <%=((rsCaseManager.Fields.Item("insId").Value == Session("insStaff_id"))?"SELECTED":"")%>><%=(rsCaseManager.Fields.Item("chvName").Value)%>
			<%	
				rsCaseManager.MoveNext();
			}
			%>
		</select></td>		
    </tr>
    <tr> 
		<td nowrap>SIN:</td>
		<td nowrap><input type="text" name="Sin" size="15" maxlength="11" tabindex="4" onChange="FormatSIN(this);"></td>
		<td nowrap>Primary Disability:</td>
		<td nowrap><select name="PrimaryDisability" tabindex="12" style="width: 150px">
			<%
			while (!rsDisability.EOF) {
			%>
				<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == 0)?"SELECTED":"")%>><%=(rsDisability.Fields.Item("chvname").Value)%>
			<%
				rsDisability.MoveNext();
			}
			%>
		</select></td>		
    </tr>
    <tr> 
		<td nowrap>PEN:</td>
		<td nowrap><input type="text" name="Pen" size="15" maxlength="9" tabindex="5"" onKeyPress="AllowNumericOnly();" ></td>
		<td nowrap>Secondary Disability:</td>
		<td nowrap><select name="SecondaryDisability" tabindex="13" style="width: 150px">
			<% 
			rsDisability.MoveFirst();
			while (!rsDisability.EOF) {
			%>
				<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == 0)?"SELECTED":"")%>><%=(rsDisability.Fields.Item("chvname").Value)%></option>
			<%
				rsDisability.MoveNext();
			}
			%>
		</select></td>		
	</tr>
	<tr> 
		<td nowrap>Gender:</td>
		<td nowrap><select name="Gender" tabindex="6">
			<option value="1">Male
			<option value="0">Female
		</select></td>
		<td nowrap>Program Standing:</td>
		<td nowrap><select name="ProgramStanding" tabindex="14" style="width: 150px">
			<option value="1">Default
			<option value="0" SELECTED>In Good Standing
		</select></td>	
	</tr>
    <tr> 
		<td nowrap>Date of Birth:</td>
		<td nowrap>
			<input type="text" name="DateOfBirth" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Past Service Received:</td>
		<td nowrap>
			<input type="checkbox" name="SetBCServed" value="1" tabindex="15" class="chkstyle">SETBC
			<input type="checkbox" name="PRCVIServed" value="1" tabindex="16" accesskey="L" class="chkstyle">PRCVI
		</td>
    </tr>
	<tr>
		<td nowrap></td>
		<td nowrap></td>
		<td nowrap>Is First Nations:</td>
		<td nowrap><input type="checkbox" name="IsFirstNations" value="1" tabindex="17" class="chkstyle">Is First Nations</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" tabindex="18" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Cancel" tabindex="19" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsRegion.Close();
rsStatus.Close();
rsDisability.Close();
rsCaseManager.Close();
%>