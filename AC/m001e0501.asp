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
	var Duration = ((Request.Form("Duration")=="")?"0":Request.Form("Duration"));
	var IsActive = ((String(Request.Form("IsActive"))=="on")?"1":"0");	
	var ProgramName = String(Request.Form("ProgramName")).replace(/'/g, "''");	
	var StudentNumber = String(Request.Form("StudentNumber")).replace(/'/g, "''");	
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");	
	var rsEducationHistory = Server.CreateObject("ADODB.Recordset");
	rsEducationHistory.ActiveConnection = MM_cnnASP02_STRING;
	rsEducationHistory.Source = "{call dbo.cp_Edu_Hstry2("+Request.Form("MM_recordId")+","+ Request.QueryString("intAdult_id") + "," + Request.Form("InstitutionName") + ",'" + ProgramName + "',"+Request.Form("ProgramType")+",'"+Request.Form("StartDate")+"','"+Request.Form("EndDate")+"',"+Duration+","+Request.Form("DurationType")+",'"+StudentNumber+"',"+IsActive+",'"+Notes+"',0,'E',0)}";
	rsEducationHistory.CursorType = 0;
	rsEducationHistory.CursorLocation = 2;
	rsEducationHistory.LockType = 3;
	rsEducationHistory.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q0501.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsDurationType = Server.CreateObject("ADODB.Recordset");
rsDurationType.ActiveConnection = MM_cnnASP02_STRING;
rsDurationType.Source = "{call dbo.cp_Duratn_Type}";
rsDurationType.CursorType = 0;
rsDurationType.CursorLocation = 2;
rsDurationType.LockType = 3;
rsDurationType.Open();

var rsEducationHistory = Server.CreateObject("ADODB.Recordset");
rsEducationHistory.ActiveConnection = MM_cnnASP02_STRING;
rsEducationHistory.Source = "{call dbo.cp_Edu_Hstry2("+Request.QueryString("intEduHst_id")+","+ Request.QueryString("intAdult_id") + ",0,'',0,'','',0,0,'',0,'',1,'Q',0)}";
rsEducationHistory.CursorType = 0;
rsEducationHistory.CursorLocation = 2;
rsEducationHistory.LockType = 3;
rsEducationHistory.Open();

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_School}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();

var rsProgramType = Server.CreateObject("ADODB.Recordset");
rsProgramType.ActiveConnection = MM_cnnASP02_STRING;
rsProgramType.Source = "{call dbo.cp_asp_lkup2(20,0,'',0,'1',0)}";
rsProgramType.CursorType = 0;
rsProgramType.CursorLocation = 2;
rsProgramType.LockType = 3;
rsProgramType.Open();
%>
<html>
<head>
	<title>Update Education Record</title>
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
				document.frm0501.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0501.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (!CheckDate(document.frm0501.StartDate.value)){
			alert("Invalid Start Date.");
			document.frm0501.StartDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0501.EndDate.value)){
			alert("Invalid End Date.");
			document.frm0501.EndDate.focus();
			return ;
		}
		document.frm0501.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0501.InstitutionName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0501">
<h5>Update Education Record</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Institution Name:</td>
		<td nowrap><select name="InstitutionName" tabindex="1" accesskey="F">
			<% 
			while (!rsInstitution.EOF) { 
			%>
				<option value="<%=(rsInstitution.Fields.Item("insSchool_id").Value)%>" <%=((rsInstitution.Fields.Item("insSchool_id").Value == rsEducationHistory.Fields.Item("insSchool_id").Value)?"SELECTED":"")%>><%=(rsInstitution.Fields.Item("chvName").Value)%></option>
			<%
				rsInstitution.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Program Name:</td>
		<td nowrap><input type="text" name="ProgramName" value="<%=(rsEducationHistory.Fields.Item("chvPgm_Name").Value)%>" maxlength="50" size="40" tabindex="2" ></td>
	</tr>
	<tr>		
		<td nowrap>Program Type:</td>
		<td nowrap><select name="ProgramType" tabindex="3">
			<% 
			while (!rsProgramType.EOF) {
			%>
				<option value="<%=(rsProgramType.Fields.Item("insProg_type_id").Value)%>" <%=((rsProgramType.Fields.Item("insProg_type_id").Value == rsEducationHistory.Fields.Item("insProg_type_id").Value)?"SELECTED":"")%>><%=(rsProgramType.Fields.Item("chvname").Value)%></option>
			<%
				rsProgramType.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input type="text" name="StartDate" value="<%=FilterDate(rsEducationHistory.Fields.Item("dtsStart").Value)%>" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
    <tr> 
		<td nowrap>End Date:</td>
		<td nowrap>
			<input type="text" name="EndDate" value="<%=FilterDate(rsEducationHistory.Fields.Item("dtsEnd_date").Value)%>" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Duration:</td>
		<td nowrap>
			<input type="text" name="Duration" value="<%=(rsEducationHistory.Fields.Item("insDurtn_Qty").Value)%>" maxlength="3" size="3" tabindex="6" onKeypress="AllowNumericOnly();" >
			<select name="DurationType" tabindex="7">
			<% 
			while (!rsDurationType.EOF) {
			%>
				<option value="<%=(rsDurationType.Fields.Item("insDuratn_type_id").Value)%>" <%=((rsDurationType.Fields.Item("insDuratn_type_id").Value == rsEducationHistory.Fields.Item("insDuratn_type_id").Value)?"SELECTED":"")%>><%=(rsDurationType.Fields.Item("chvDuratn_desc").Value)%></option>
			<%
				rsDurationType.MoveNext();
			}
			%>
			</select>
		</td>
	</tr>
    <tr> 
		<td nowrap>Student Number:</td>
		<td nowrap><input type="text" name="StudentNumber" value="<%=(rsEducationHistory.Fields.Item("chvStdnum").Value)%>" maxlength="15" size="15" tabindex="8"></td>
    </tr>
    <tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" <%=((rsEducationHistory.Fields.Item("bitIs_Active").Value=="1")?"CHECKED":"")%> tabindex="9" class="chkstyle"></td>
    </tr>	
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="10" accesskey="L"><%=(rsEducationHistory.Fields.Item("chvnotes").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="11" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="12" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="13" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsEducationHistory.Fields.Item("intEduHst_id").Value %>">
</form>
</body>
</html>
<%
rsDurationType.Close();
rsEducationHistory.Close();
rsInstitution.Close();
rsProgramType.Close();
%>