<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var FirstName = String(Request.Form("FirstName")).replace(/'/g, "''");	
	var MiddleName = String(Request.Form("MiddleName")).replace(/'/g, "''");	
	var LastName = String(Request.Form("LastName")).replace(/'/g, "''");
	var IsFirstNation = ((Request.Form("IsFirstNation")=="1")?"1":"0");	
	var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
	rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
	rsPILATStudent.Source = "{call dbo.cp_pilat_student("+ Request.QueryString("intPStdnt_id") + ",'"+FirstName+"','"+MiddleName+"','"+LastName+"','"+Request.Form("Sin")+"','"+Request.Form("DateOfBirth")+"',"+Request.Form("Status")+","+Request.Form("Disability")+","+Request.Form("CaseManager")+","+Request.Form("Region")+",'"+Request.Form("Pen")+"',"+IsFirstNation+","+Request.Form("Gender")+","+Session("insStaff_id")+",1,0,'',0,'E',0)}";
	rsPILATStudent.CursorType = 0;
	rsPILATStudent.CursorLocation = 2;
	rsPILATStudent.LockType = 3;
	rsPILATStudent.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m022e0101.asp&intPStdnt_id="+Request.QueryString("intPStdnt_id"));
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
rsStatus.Source = "{call dbo.cp_StdStatus}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();

var rsDisability = Server.CreateObject("ADODB.Recordset");
rsDisability.ActiveConnection = MM_cnnASP02_STRING;
rsDisability.Source = "{call dbo.cp_StdDsbty}";
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

var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStudent.Source = "{call dbo.cp_pilat_student("+ Request.QueryString("intPStdnt_id") + ",'','','','','',0,0,0,0,'',0,0,0,1,0,'',1,'Q',0)}";
rsPILATStudent.CursorType = 0;
rsPILATStudent.CursorLocation = 2;
rsPILATStudent.LockType = 3;
rsPILATStudent.Open();
%>								
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>General Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0101.reset();
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
		document.frm0101.Sin.value = LeaveDigits(document.frm0101.Sin.value);
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0101.FirstName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>First Name:</td>
		<td nowrap width="150"><input type="text" name="FirstName" value="<%=Trim(rsPILATStudent.Fields.Item("chvFst_name").Value)%>" tabindex="1" accesskey="F"></td>	 
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="7">
			<%
			while (!rsRegion.EOF) {
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsRegion.Fields.Item("insRegion_num").Value == rsPILATStudent.Fields.Item("insRegion_num").Value)?"SELECTED":"")%> ><%=(rsRegion.Fields.Item("chvname").Value)%></option>
			<%
				rsRegion.MoveNext();
			}
			%>
		</select></td>		
	</tr>
	<tr> 
		<td nowrap>Middle Name:</td>
		<td nowrap><input type="text" name="MiddleName" value="<%=Trim(rsPILATStudent.Fields.Item("chvMdl_name").Value)%>" maxlength="50" tabindex="2"></td>	
		<td nowrap>Case Manager:</td>
		<td nowrap><select name="CaseManager" tabindex="8" style="width: 200px">
			<%
			while (!rsCaseManager.EOF) {
			%>
				<option value="<%=(rsCaseManager.Fields.Item("insId").Value)%>" <%=((rsCaseManager.Fields.Item("insId").Value == rsPILATStudent.Fields.Item("insCase_mngr_id").Value)?"SELECTED":"")%> ><%=(rsCaseManager.Fields.Item("chvName").Value)%></option>
			<%
				rsCaseManager.MoveNext();
			}
			%>
		</select></td>				
	</tr>
	<tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" value="<%=Trim(rsPILATStudent.Fields.Item("chvLst_name").Value)%>" maxlength="50" tabindex="3"></td>	
		<td nowrap>Disability:</td>
		<td nowrap><select name="Disability" tabindex="9" style="width: 200px">
			<%
			while (!rsDisability.EOF) {
			%>
				<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == rsPILATStudent.Fields.Item("insDsbty1_id").Value)?"SELECTED":"")%> ><%=(rsDisability.Fields.Item("chvname").Value)%></option>
			<%
				rsDisability.MoveNext();
			}
			%>
		</select></td>				
	</tr>
    <tr> 
		<td nowrap>SIN:</td>
		<td nowrap><input type="text" name="Sin" value="<%=FormatSIN(Trim(rsPILATStudent.Fields.Item("chrSIN_no").Value))%>" size="15" maxlength="11" tabindex="4" onChange="FormatSIN(this);" ></td>
		<td nowrap>Date of Birth:</td>
		<td nowrap><input type="text" name="DateOfBirth" value="<%=FilterDate(rsPILATStudent.Fields.Item("dtsBirth_date").Value)%>" size="11" maxlength="10" tabindex="10" onChange="FormatDate(this)"><span style="font-size: 7pt">(mm/dd/yyyy)</span></td>
	</tr>
    <tr> 
		<td nowrap>PEN:</td>
		<td nowrap><input type="text" name="Pen" value="<%=Trim(rsPILATStudent.Fields.Item("chrPEN_num").Value)%>" size="9" maxlength="9" tabindex="5" onKeypress="AllowNumericOnly();" ></td>
		<td nowrap>Is First Nation:</td>
		<td nowrap><input type="checkbox" name="FirstNation" tabindex="11" accesskey="L" <%=((rsPILATStudent.Fields.Item("bitIs_FirstNations").Value == "1")?"CHECKED":"")%> class="chkstyle"></td>
	</tr>
	<tr>
		<td nowrap>Gender:</td>
		<td nowrap><select name="Gender" tabindex="6">
			<option value="1" <%=((rsPILATStudent.Fields.Item("bitGender_is_male").Value == 1)?"Selected":"")%>>Male
			<option value="0" <%=((rsPILATStudent.Fields.Item("bitGender_is_male").Value == 0)?"Selected":"")%>>Female
		</select></td>
		<td colspan="2"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="12" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="13" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=(rsPILATStudent.Fields.Item("intPStdnt_id").Value)%>">
<input type="hidden" name="Status" value="<%=(rsPILATStudent.Fields.Item("insStdnt_Status_id").Value)%>">
</form>
</body>
</html>
<%
rsRegion.Close();
rsStatus.Close();
rsDisability.Close();
rsCaseManager.Close();
rsPILATStudent.Close();
%>