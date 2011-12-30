<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");	
	var MiddleName = String(Request.Form("MiddleName")).replace(/'/g, "''");	
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");
	var IsFirstNation = ((Request.Form("IsFirstNation")=="1")?"1":"0");	
	var cmdInsertPILATStudent = Server.CreateObject("ADODB.Command");
	cmdInsertPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertPILATStudent.CommandText = "dbo.cp_PILAT_Student";
	cmdInsertPILATStudent.CommandType = 4;
	cmdInsertPILATStudent.CommandTimeout = 0;
	cmdInsertPILATStudent.Prepared = true;
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@intpID", 3, 1,1,0));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chvFst_name", 200, 1,20,FirstName));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chvMdl_name", 200, 1,20,MiddleName));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chvLst_name", 200, 1,20,LastName));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chrSIN_no", 129, 1,20,Request.Form("SIN")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@dtsBirth_date", 200,1,30,Request.Form("DateOfBirth")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insStdnt_Status_id", 2, 1,1,Request.Form("Status")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insDsbty1_id", 2, 1,1,Request.Form("Disability")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insCase_mngr_id", 2, 1,1,Request.Form("CaseManager")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insRegion_num", 2, 1,1,Request.Form("Region")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chrPEN_num", 129, 1,20,Request.Form("PEN")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@bitIs_FirstNations", 2, 1,1,IsFirstNation));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@bitGender_is_male", 2, 1,1,Request.Form("Gender")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@inspSrtBy", 2, 1,1,1));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chvFilter", 200, 1,1,""));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertPILATStudent.Parameters.Append(cmdInsertPILATStudent.CreateParameter("@intRtnValue", 2, 2));	
	cmdInsertPILATStudent.Execute();
	Response.Redirect("m022FS3.asp?intPStdnt_id="+cmdInsertPILATStudent.Parameters.Item("@intRtnValue").Value);
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

var FirstName = "";
var MiddleName = "";
var LastName = "";
var Region = "";
var Disability = 0;
var IsFirstNation = 0;
var DateOfBirth = "";
var Pen = "";
var Sin = "";
var Status = 0;
var CaseManager = 0;
var Gender = 1;

if (String(Request.QueryString("IsNew"))=="No") {
	var rsClient = Server.CreateObject("ADODB.Recordset");
	rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
	rsClient.CursorType = 0;
	rsClient.CursorLocation = 2;
	rsClient.LockType = 3;
	rsClient.Open();
	FirstName=rsClient.Fields.Item("chvFst_Name").Value;
	MiddleName=rsClient.Fields.Item("chvMdl_name").Value;
	LastName=rsClient.Fields.Item("chvLst_Name").Value;
	Status=rsClient.Fields.Item("insStdnt_status_id").Value;
	Region=rsClient.Fields.Item("insRegion_num").Value;
	Disability=rsClient.Fields.Item("insDsbty1_id").Value;
	CaseManager=rsClient.Fields.Item("insCase_mngr_id").Value;
	Gender=rsClient.Fields.Item("bitGender_is_male").Value;
	IsFirstNation=rsClient.Fields.Item("bitIs_FirstNations").Value;
	Pen=rsClient.Fields.Item("chvPENno").Value;
	Sin=rsClient.Fields.Item("chrSIN_no").Value;
	DateOfBirth=FilterDate(rsClient.Fields.Item("dtsBirth_date").Value);
}
%>
<html>
<head>
	<title>New Temp Student</title>
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
	function InputDate(DateField){
		DateField.value = FormatDate(LeaveDigits(DateField));
	}
		
	function Save(){
		if (!CheckSIN(document.frm0102.Sin.value)) {
			alert("Invalid Social Insurance Number.");
			document.frm0102.Sin.focus();
			return ;
		}
		if (!CheckDate(document.frm0102.DateOfBirth.value)){
			alert("Invalid Date of Birth.");
			document.frm0102.DateOfBirth.focus();
			return ;
		}
		if (Trim(document.frm0102.FirstName.value)=="") {
			alert("Enter First Name.");
			document.frm0102.FirstName.focus();
			return ;
		}
		if (Trim(document.frm0102.LastName.value)=="") {
			alert("Enter Last Name.");
			document.frm0102.LastName.focus();
			return ;
		}
		document.frm0102.Sin.value = LeaveDigits(document.frm0102.Sin.value);
		document.frm0102.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0102.FirstName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<h5>New PILAT Student:</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>First Name:</td>
		<td width="200"><input type="text" name="FirstName" value="<%=FirstName%>" maxlength="20" tabindex="1" accesskey="F"></td>
		<td nowrap>Region:</td>
		<td width="200"><select name="Region" tabindex="7" style="width: 200px">
			<% 
			while (!rsRegion.EOF) {
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsRegion.Fields.Item("insRegion_num").Value == Region)?"SELECTED":"")%>><%=(rsRegion.Fields.Item("chvname").Value)%></option>
			<%
				rsRegion.MoveNext();
			}
			%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Middle Name:</td>
		<td nowrap><input type="text" name="MiddleName" value="<%=MiddleName%>" maxlength="20" tabindex="2"></td>
		<td nowrap>Case Manager:</td>
		<td nowrap><select name="CaseManager" tabindex="8" style="width: 200px">
		<% 
		while (!rsCaseManager.EOF) {
		%>
			<option value="<%=(rsCaseManager.Fields.Item("insId").Value)%>" <%=((rsCaseManager.Fields.Item("insId").Value == CaseManager)?"SELECTED":"")%>><%=(rsCaseManager.Fields.Item("chvName").Value)%></option>
		<%	
			rsCaseManager.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" value="<%=LastName%>" maxlength="20" tabindex="3"></td>
		<td nowrap>Disability:</td>
		<td nowrap><select name="Disability" tabindex="9" style="width: 200px" accesskey="D">
			<% 
			while (!rsDisability.EOF) {
			%>
				<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == Disability)?"SELECTED":"")%>><%=(rsDisability.Fields.Item("chvname").Value)%></option>
			<%
				rsDisability.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>SIN:</td>
		<td nowrap><input type="text" name="Sin" value="<%=Sin%>" size="15" maxlength="11" tabindex="4" onChange="FormatSIN(this);"></td>
		<td nowrap>Date of Birth:</td>
		<td nowrap>
			<input type="text" name="DateOfBirth" value="<%=DateOfBirth%>" size="11" maxlength="10" tabindex="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>PEN:</td>
		<td nowrap><input type="text" name="Pen" value="<%=Pen%>" size="15" maxlength="9" tabindex="5" onKeyPress="AllowNumericOnly();"></td>
		<td nowrap>Is First Nation:</td>
		<td nowrap><input type="checkbox" name="IsFirstNation" value="1" <%=((IsFirstNation==1)?"CHECKED":"")%> tabindex="11" accesskey="L" class="chkstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Gender:</td>
		<td nowrap><select name="Gender" tabindex="6">
			<option value="1" <%=((Gender==1)?"SELECTED":"")%>>Male 
			<option value="0" <%=((Gender==0)?"SELECTED":"")%>>Female 
		</select></td>
		<td colspan="2"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="12" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Cancel" tabindex="13" onClick="window.close();" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
<input type="hidden" name="Status" value="<%=Status%>">
</form>
</body>
</html>
<%
rsRegion.Close();
rsStatus.Close();
rsDisability.Close();
rsCaseManager.Close();
%>