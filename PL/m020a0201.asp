<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var rsIssueNote = Server.CreateObject("ADODB.Recordset");
	rsIssueNote.ActiveConnection = MM_cnnASP02_STRING;
	rsIssueNote.Source = "{call dbo.cp_pjt_issues(0,'"+String(Request.Form("Subject")).replace(/'/g, "''")+"','"+String(Request.Form("Notes")).replace(/'/g, "''")+"',"+Request.Form("Function")+","+Session("insStaff_id")+","+Request.Form("Priority")+","+Request.Form("Status")+",'"+String(Request.Form("Version")).replace(/'/g, "''")+"',0,0,"+Request.Form("AssignedTo")+",'" + CurrentDate() + "',1,0,'',0,'A',0)}";
	rsIssueNote.CursorType = 0;
	rsIssueNote.CursorLocation = 2;
	rsIssueNote.LockType = 3;
	rsIssueNote.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_ASP_Lkup(702)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_Lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsPriority = Server.CreateObject("ADODB.Recordset");
rsPriority.ActiveConnection = MM_cnnASP02_STRING;
rsPriority.Source = "{call dbo.cp_pjt_priorities(0,'',0,'',0,'Q',0)}";
rsPriority.CursorType = 0;
rsPriority.CursorLocation = 2;
rsPriority.LockType = 3;
rsPriority.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_pjt_statues(0,'',0,'Q',0)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();
%>
<html>
<head>
	<title>New Issue</title>
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
	function AlertTimeOut(){
		alert("Session will timeout in 5 minutes.  Please save your work.");		
	}
	
	function Save(){
		if (Trim(document.frm0201.Subject.value)=="") {
			if (!confirm("Save without Subject?")) return ;
		}		
		if (!CheckTextArea(document.frm0201.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
		document.frm0201.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0201.Subject.focus();setTimeout('AlertTimeOut()',2100000);">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>New Issue</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><input type="text" name="Subject" maxlength="100" size="80" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Function:</td>
		<td nowrap><select name="Function" tabindex="2">
			<% 
			while (!rsFunction.EOF) { 
			%>
				<option value="<%=(rsFunction.Fields.Item("insFTNid").Value)%>"><%=(rsFunction.Fields.Item("intMODno").Value)%>.<%=(rsFunction.Fields.Item("insFSTno").Value)%>.<%=(rsFunction.Fields.Item("insFSTsubno").Value)%>.:<%=(rsFunction.Fields.Item("ncvFTNname").Value)%> (<%=rsFunction.Fields.Item("ncvMODname").Value%>)</option>
			<%
				rsFunction.MoveNext();
			}
			%>
		</select></td>
    </tr>	
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap><textarea name="Notes" cols="80" rows="18" tabindex="3"></textarea></td>
    </tr>
	<tr>
		<td nowrap>Priority:</td>
		<td nowrap><select name="Priority" tabindex="4">
			<% 
			while (!rsPriority.EOF) { 
			%>
				<option value="<%=(rsPriority.Fields.Item("intPriority_id").Value)%>" <%=((rsPriority.Fields.Item("intPriority_id").Value=="3")?"SELECTED":"")%>><%=(rsPriority.Fields.Item("ncvPriority_desc").Value)%></option>
			<% 
				rsPriority.MoveNext();
			}
			%>
		</select></td>
	</tr>	
	<tr>
		<td nowrap>Status:</td>
		<td nowrap><select name="Status" tabindex="5">
			<% 
			while (!rsStatus.EOF) { 
			%>
				<option value="<%=(rsStatus.Fields.Item("intStatus_id").Value)%>"><%=(rsStatus.Fields.Item("ncvStatus").Value)%></option>
			<% 
				rsStatus.MoveNext();
			}
			%>
		</select></td>
	</tr>	
	<tr>
		<td nowrap>Version:</td>
		<td nowrap><input type="text" name="Version" size="8" tabindex="6"></td>
	</tr>	
	<tr>
		<td nowrap>Assigned To:</td>
		<td nowrap><select name="AssignedTo" tabindex="7" accesskey="L">
			<% 
			while (!rsStaff.EOF) { 
			%>
				<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value=="135")?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvname").Value)%></option>
			<% 
				rsStaff.MoveNext();
			}
			%>
        </select></td>
    </tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="8" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="9" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsFunction.Close();
rsStaff.Close();
rsPriority.Close();
rsStatus.Close();
%>