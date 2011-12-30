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
	var Duration = ((Request.Form("Duration")=="")?"0":Request.Form("Duration"));
	var ProgramName = String(Request.Form("ProgramName")).replace(/'/g, "''");	
	var StudentNumber = String(Request.Form("StudentNumber")).replace(/'/g, "''");	
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");	
	var rsEducationHistory = Server.CreateObject("ADODB.Recordset");
	rsEducationHistory.ActiveConnection = MM_cnnASP02_STRING;
	rsEducationHistory.Source = "{call dbo.cp_Edu_Hstry2(0,"+ Request.QueryString("intAdult_id") + "," + Request.Form("InstitutionName") + ",'" + ProgramName + "',"+Request.Form("ProgramType")+",'"+Request.Form("StartDate")+"','"+Request.Form("EndDate")+"',"+Duration+","+Request.Form("DurationType")+",'"+StudentNumber+"',1,'"+Notes+"',0,'A',0)}";
	rsEducationHistory.CursorType = 0;
	rsEducationHistory.CursorLocation = 2;
	rsEducationHistory.LockType = 3;
	rsEducationHistory.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsDurationType = Server.CreateObject("ADODB.Recordset");
rsDurationType.ActiveConnection = MM_cnnASP02_STRING;
rsDurationType.Source = "{call dbo.cp_Duratn_Type}";
rsDurationType.CursorType = 0;
rsDurationType.CursorLocation = 2;
rsDurationType.LockType = 3;
rsDurationType.Open();

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_School}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsProgramType = Server.CreateObject("ADODB.Recordset");
rsProgramType.ActiveConnection = MM_cnnASP02_STRING;
rsProgramType.Source = "{call dbo.cp_ASP_lkup2(20,0,'',0,'1',0)}";
rsProgramType.CursorType = 0;
rsProgramType.CursorLocation = 2;
rsProgramType.LockType = 3;
rsProgramType.Open();
%>
<html>
<head>
	<title>New Education Record For <%=rsClient.Fields.Item("chvName").Value%></title>
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
		document.frm0501.btnSave.disabled = true;
	}
	</script>
</head>
<body onLoad="javascript:document.frm0501.InstitutionName.focus()">
<form name="frm0501" method="POST" action="<%=MM_editAction%>">
<h5>New Education Record:</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Institution Name:</td>
		<td nowrap><select name="InstitutionName" tabindex="1" accesskey="F">
		<%
		while (!rsInstitution.EOF) { 
		%>
			<option value="<%=(rsInstitution.Fields.Item("insSchool_id").Value)%>"><%=(rsInstitution.Fields.Item("chvName").Value)%></option>
		<%
			rsInstitution.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Program Name:</td>
		<td nowrap><input type="text" name="ProgramName" maxlength="50" tabindex="2" size="40"></td>
	</tr>
	<tr>
		<td nowrap>Program Type:</td>
		<td nowrap><select name="ProgramType" tabindex="3">
		<% 
		while (!rsProgramType.EOF) {
		%>
			<option value="<%=(rsProgramType.Fields.Item("insProg_type_id").Value)%>"><%=(rsProgramType.Fields.Item("chvname").Value)%></option>
		<%
			rsProgramType.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input type="text" name="StartDate" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this);">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>End Date:</td>
		<td nowrap>
			<input type="text" name="EndDate" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this);">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Duration:</td>
		<td nowrap>
			<input type="text" name="Duration" maxlength="3" size="3" tabindex="6" onKeypress="AllowNumericOnly();">
			<select name="DurationType" tabindex="7">
			<% 
			while (!rsDurationType.EOF) {
			%>
				<option value="<%=(rsDurationType.Fields.Item("insDuratn_type_id").Value)%>" <%=((rsDurationType.Fields.Item("insDuratn_type_id").Value=="5")?"SELECTED":"")%>><%=(rsDurationType.Fields.Item("chvDuratn_desc").Value)%>
			<%
				rsDurationType.MoveNext();
			}
			%>
			</select>
		</td>
    </tr>
    <tr> 
		<td nowrap>Student Number:</td>
		<td nowrap><input type="text" name="StudentNumber" size="15" maxlength="15" tabindex="8"></td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="9" accesskey="L"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" tabindex="10" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="11" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsDurationType.Close();
rsInstitution.Close();
rsClient.Close();
rsProgramType.Close();
%>