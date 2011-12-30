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
	rsEmployment.Source = "{call dbo.cp_EmplyInfo2("+Request.Form("MM_recordId")+",0,"+Request.QueryString("insStaff_id")+",1,"+Request.Form("CompanyName")+",0,'"+Request.Form("EmploymentType")+"','"+Request.Form("EmploymentDuration")+"','"+Request.Form("StartDate")+"','"+Request.Form("EndDate")+"','"+JobDescription.replace(/'/g, "''")+"',0,'E',0)}";
	rsEmployment.CursorType = 0;
	rsEmployment.CursorLocation = 2;
	rsEmployment.LockType = 3;
	rsEmployment.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m002q0102.asp&insStaff_id="+Request.QueryString("insStaff_id"))
}

var rsEmployment = Server.CreateObject("ADODB.Recordset");
rsEmployment.ActiveConnection = MM_cnnASP02_STRING;
rsEmployment.Source = "{call dbo.cp_EmplyInfo2("+ Request.QueryString("intEmply_id") + ",0,0,0,0,0,'','','','','',1,'Q',0)}";
rsEmployment.CursorType = 0;
rsEmployment.CursorLocation = 2;
rsEmployment.LockType = 3;
rsEmployment.Open();

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
<body onLoad="javascript:document.frm0102.CompanyName.focus()" >
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Company Name:</td>
		<td nowrap><select name="CompanyName" tabindex="1" accesskey="F">
		<% 
		while (!rsCompany.EOF) 
		{
		%>
			<option value="<%=(rsCompany.Fields.Item("intCompany_id").Value)%>" <%=((rsCompany.Fields.Item("intCompany_id").Value == rsEmployment.Fields.Item("intCompany_id").Value)?"SELECTED":"")%> ><%=(rsCompany.Fields.Item("chvCompany_Name").Value)%></option>
		<%
			rsCompany.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Employment Type:</td>
		<td nowrap><select name="EmploymentType" tabindex="2">
	  		<option value="R" <%=((rsEmployment.Fields.Item("chrEmploy_Type").Value == "R")?"SELECTED":"")%>>Regular
			<option value="S" <%=((rsEmployment.Fields.Item("chrEmploy_Type").Value == "S")?"SELECTED":"")%>>Self Employed
			<option value="P" <%=((rsEmployment.Fields.Item("chrEmploy_Type").Value == "P")?"SELECTED":"")%>>PSTP
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Duration:</td>
		<td nowrap><select name="EmploymentDuration" tabindex="3">
			<option value="T" <%=((rsEmployment.Fields.Item("chrEmploy_Dur").Value == "T")?"SELECTED":"")%>>Temporary
			<option value="P" <%=((rsEmployment.Fields.Item("chrEmploy_Dur").Value == "P")?"SELECTED":"")%>>Permanent
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input type="text" name="StartDate" value="<%=FilterDate(rsEmployment.Fields.Item("dtmDate_from").Value)%>" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>End Date:</td>
		<td nowrap>
			<input type="text" name="EndDate" value="<%=FilterDate(rsEmployment.Fields.Item("dtmDate_To").Value)%>" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>		
    <tr> 
		<td nowrap valign="top">Job Description:</td>
		<td nowrap><textarea name="JobDescription" accesskey="L" rows="5" cols="65" tabindex="6"><%=(rsEmployment.Fields.Item("chvDuties").Value)%></textarea></td>
    </tr>
</table>	
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>
			<input type="button" value="Save" tabindex="7" onClick="Save();" class="btnstyle">&nbsp;&nbsp;
			<input type="reset" value="Undo Changes" tabindex="8" class="btnstyle">&nbsp;&nbsp;
			<input type="button" value="Close" tabindex="9" onClick="history.back();" class="btnstyle">
		</td>
	</tr>
</table>  
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsEmployment.Fields.Item("intEmply_id").Value %>">
</form>
</body>
</html>
<%
rsEmployment.Close();
rsCompany.Close();
%>