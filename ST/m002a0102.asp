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
	var JobDescription = String(Request.Form("JobDescription"))
	var rsEmployment = Server.CreateObject("ADODB.Recordset");
	rsEmployment.ActiveConnection = MM_cnnASP02_STRING;
	rsEmployment.Source = "{call dbo.cp_EmplyInfo2(0,0,"+Request.Form("insStaff_id")+",1,"+Request.Form("OrganizationName")+",0,'"+Request.Form("EmploymentType")+"','"+Request.Form("EmploymentDuration")+"','"+Request.Form("StartDate")+"','"+Request.Form("EndDate")+"','"+JobDescription.replace(/'/g, "''")+"',0,'A',0)}";
	rsEmployment.CursorType = 0;
	rsEmployment.CursorLocation = 2;
	rsEmployment.LockType = 3;
	rsEmployment.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,1,0,'',0,'Q',0)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();
%>
<html>
<head>
	<title>New Employment Record</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="javascript" src="../js/MyFunctions.js"></script>
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
	<script language="javascript">
	function Save(){
		if (!CheckDate(document.frm0102.StartDate.value)){
			alert("Invalid Start Date.");
			document.frm0102.StartDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0102.EndDate.value)){
			alert("Invalid End Date.");
			document.frm0102.EndDate.focus();
			return ;
		}
		document.frm0102.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0102.OrganizationName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<h5>New Employment Record:</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Organization Name:</td>
		<td nowrap><select name="OrganizationName" tabindex="1" accesskey="F">
		<% 
		while (!rsCompany.EOF) {
		%>
			<option value="<%=(rsCompany.Fields.Item("intCompany_id").Value)%>" <%=((rsCompany.Fields.Item("intCompany_id").Value == 0)?"SELECTED":"")%>><%=(rsCompany.Fields.Item("chvCompany_Name").Value)%></option>
		<%
			rsCompany.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Employment Type:</td>
		<td nowrap><select name="EmploymentType" tabindex="2">
			<option value="R">Regular 
			<option value="S">Self Employed 
			<option value="P">PSTP 
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Duration:</td>
		<td nowrap><select name="EmploymentDuration" tabindex="3">
			<option value="T">Temporary 
			<option value="P">Permanent 
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Start Date:</td>
		<td nowrap> 
			<input type="text" name="StartDate" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
    <tr> 
		<td nowrap>End Date:</td>
		<td nowrap> 
			<input type="text" name="EndDate" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span> 
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Job Description:</td>
		<td nowrap valign="top"><textarea name="JobDescription" rows="5" cols="65" tabindex="6" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="7" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="8" onClick="window.close();" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
<input type="hidden" name="insStaff_id" value="<%=Request.QueryString("insStaff_id")%>">
</form>
</body>
</html>
<%
rsCompany.Close();
%>