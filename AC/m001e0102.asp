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
	var JobDescription = String(Request.Form("JobDescription"))
	var rsEmployment = Server.CreateObject("ADODB.Recordset");
	rsEmployment.ActiveConnection = MM_cnnASP02_STRING;
	rsEmployment.Source = "{call dbo.cp_EmplyInfo2("+Request.Form("MM_recordId")+","+ Request.QueryString("intAdult_id") + ",0,0,"+Request.Form("CompanyName")+",0,'"+Request.Form("EmploymentType")+"','"+Request.Form("EmploymentDuration")+"','"+Request.Form("StartDate")+"','"+Request.Form("EndDate")+"','"+JobDescription.replace(/'/g, "''")+"',0,'E',0)}";
	rsEmployment.CursorType = 0;
	rsEmployment.CursorLocation = 2;
	rsEmployment.LockType = 3;
//	Response.Redirect(rsEmployment.Source);
	rsEmployment.Open();	
	Response.Redirect("UpdateSuccessful.asp?page=m001q0102.asp&intAdult_id="+Request.QueryString("intAdult_id"))
}

var rsEmployment = Server.CreateObject("ADODB.Recordset");
rsEmployment.ActiveConnection = MM_cnnASP02_STRING;
rsEmployment.Source = "{call dbo.cp_EmplyInfo2("+ Request.QueryString("intEmply_id") + ",0,0,0,0,0,'','','','','',1,'Q',0)}";
rsEmployment.CursorType = 0;
rsEmployment.CursorLocation = 2;
rsEmployment.LockType = 3;
rsEmployment.Open();

var rsEmploymentType = Server.CreateObject("ADODB.Recordset");
rsEmploymentType.ActiveConnection = MM_cnnASP02_STRING;
rsEmploymentType.Source = "{call dbo.cp_employ_type(0,'','',0,'Q',0)}"
rsEmploymentType.CursorType = 0;
rsEmploymentType.CursorLocation = 2;
rsEmploymentType.LockType = 3;
rsEmploymentType.Open();

var rsEmployCompany = Server.CreateObject("ADODB.Recordset");
rsEmployCompany.ActiveConnection = MM_cnnASP02_STRING;
rsEmployCompany.Source = "{call dbo.cp_Company2("+rsEmployment.Fields.Item("intCompany_id").Value+",'',0,0,0,0,0,1,0,'',1,'Q',0)}";
rsEmployCompany.CursorType = 0;
rsEmployCompany.CursorLocation = 2;
rsEmployCompany.LockType = 3;
rsEmployCompany.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsWorkType = Server.CreateObject("ADODB.Recordset");
rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
rsWorkType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
rsWorkType.CursorType = 0;
rsWorkType.CursorLocation = 2;
rsWorkType.LockType = 3;
rsWorkType.Open();

var WorkType = ((String(Request.Form("WorkType"))=="undefined")?rsEmployCompany.Fields.Item("insWork_Typ_id").Value:Request.Form("WorkType"));

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_get_company_work_type("+WorkType+",0)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();
%>
<html>
<head>
	<title>Update Employment Information</title>
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
			case 85:
				//alert("U");
				document.frm0102.reset();
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
		if (document.frm0102.CompanyName.value <= 0) {
			alert("Select a company.");
			document.frm0102.CompanyName.focus();
			return;
		}
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
		document.frm0102.MM_update.value = "true";
		document.frm0102.submit();
	}
	</script>	
</head>
<body onLoad="javascript:document.frm0102.CompanyName.focus()" >
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Work Type:</td>
		<td nowrap><select name="WorkType" tabindex="1" onChange="document.frm0102.submit();" accesskey="F">
		<% 
		while (!rsWorkType.EOF) {
		%>
			<option value="<%=(rsWorkType.Fields.Item("intWork_type_id").Value)%>" <%=((rsWorkType.Fields.Item("intWork_type_id").Value==rsEmployCompany.Fields.Item("insWork_Typ_id").Value)?"SELECTED":"")%> <%=((rsWorkType.Fields.Item("intWork_type_id").Value==Request.Form("WorkType"))?"SELECTED":"")%>><%=(rsWorkType.Fields.Item("chvWork_type_desc").Value)%></option>
		<%
			rsWorkType.MoveNext();
		}
		%>
        </select></td>
    </tr>
	<tr> 
		<td nowrap>Company Name:</td>
		<td nowrap><select name="CompanyName" tabindex="2">
		<% 
		while (!rsCompany.EOF) {
		%>
			<option value="<%=(rsCompany.Fields.Item("intCompany_id").Value)%>" <%=((rsCompany.Fields.Item("intCompany_id").Value == rsEmployment.Fields.Item("intCompany_id").Value)?"SELECTED":"")%> ><%=(rsCompany.Fields.Item("chvOrg_Name").Value)%></option>
		<%
			rsCompany.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Employment Type:</td>
		<td nowrap><select name="EmploymentType" tabindex="3">
		<%
		while (!rsEmploymentType.EOF) {
		%>
			<option value="<%=rsEmploymentType.Fields.Item("chrEmploy_Type").Value%>" <%=((rsEmploymentType.Fields.Item("chrEmploy_Type").Value == rsEmployment.Fields.Item("chrEmploy_Type").Value)?"SELECTED":"")%>><%=rsEmploymentType.Fields.Item("chvEmploy_Desc").Value%>
		<%
		rsEmploymentType.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Duration:</td>
		<td nowrap><select name="EmploymentDuration" tabindex="4">
			<option value="">N/A
			<option value="T" <%=((rsEmployment.Fields.Item("chrEmploy_Dur").Value == "T")?"SELECTED":"")%>>Temporary
			<option value="P" <%=((rsEmployment.Fields.Item("chrEmploy_Dur").Value == "P")?"SELECTED":"")%>>Permanent
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input name="StartDate" type="text" value="<%=FilterDate(rsEmployment.Fields.Item("dtmDate_from").Value)%>" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>End Date:</td>
		<td nowrap>
			<input name="EndDate" type="text"value="<%=FilterDate(rsEmployment.Fields.Item("dtmDate_To").Value)%>" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>		
    <tr> 
		<td nowrap valign="top">Job Description:</td>
		<td nowrap valign="top"><textarea name="JobDescription" accesskey="L" rows="5" cols="65" tabindex="7"><%=(rsEmployment.Fields.Item("chvDuties").Value)%></textarea></td>
    </tr>
</table>	
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="8" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="9" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="10" onClick="history.back();" class="btnstyle"></td>
	</tr>
</table>  
<input type="hidden" name="MM_update" value="false">
<input type="hidden" name="MM_recordId" value="<%= rsEmployment.Fields.Item("intEmply_id").Value %>">
</form>
</body>
</html>
<%
rsEmployment.Close();
rsClient.Close();
rsCompany.Close();
%>
