<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_Insert"))=="true"){
	var IssueResponse = String(Request.Form("Response")).replace(/'/g, "''");	
	var Version = String(Request.Form("Version")).replace(/'/g, "''");	
	var IsApproved = ((Request.Form("Approved")=="on") ? "1":"0");
	var IsTested = ((Request.Form("Tested")=="on") ? "1":"0");
	var rsIssue = Server.CreateObject("ADODB.Recordset");
	rsIssue.ActiveConnection = MM_cnnASP02_STRING;
	rsIssue.Source = "{call dbo.cp_pjt_responses("+Request.QueryString("intIssue_id")+",'"+IssueResponse+"',0,"+Session("insStaff_id")+","+Request.Form("Priority")+","+Request.Form("CurrentStatus")+",'"+Version+"',"+IsApproved+","+IsTested+","+Request.Form("AssignedTo")+",'"+CurrentDate()+"',"+Request.Form("FunctionID")+",'',0,'A',0)}";
	rsIssue.CursorType = 0;
	rsIssue.CursorLocation = 2;
	rsIssue.LockType = 3;
	rsIssue.Open();
	Response.Redirect("m020q0201.asp");
}

var rsIssue = Server.CreateObject("ADODB.Recordset");
rsIssue.ActiveConnection = MM_cnnASP02_STRING;
rsIssue.Source = "{call dbo.cp_pjt_issues("+Request.QueryString("intIssue_id")+",'','',0,0,0,0,'',0,0,0,'',1,0,'',1,'Q',0)}";
rsIssue.CursorType = 0;
rsIssue.CursorLocation = 2;
rsIssue.LockType = 3;
rsIssue.Open();

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
	<title>Issue Response</title>
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
				document.frm0201.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function AlertTimeOut(){
		alert("Session will timeout in 5 minutes.  Please save your work.");		
	}
		
	function Save(){
		if (document.frm0201.Response.value.length> 4000) {
			alert("Response has exceeded 4000 characters limit.");
			return ;			
		}	
		document.frm0201.submit();
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=600,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	
	function Init(){
		document.frm0201.Response.focus();
		setTimeout('AlertTimeOut()',2100000);	
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>Issue Response</h5>
<table cellpadding="2" cellspacing="1" width="60%">
    <tr> 
		<td class="headrow" align="left">Subject:</td>
		<td colspan="2" style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvSubject").Value)%>&nbsp;</td>
		<td><input type="button" value="History" onClick="openWindow('m020e0202.asp?intIssue_id=<%=Request.QueryString("intIssue_id")%>','');" class="btnstyle"></td>
    </tr>
    <tr> 
		<td class="headrow" align="left">Function:</td>
		<td colspan="3" style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("ncvFTNname").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" align="left">Created By:</td>
		<td colspan="3" style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvSubmitted_by").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" align="left">Date:</td>
		<td style="border: solid 1px #CCCCCC" width="170"><%=(rsIssue.Fields.Item("dtsDate_submitted").Value)%>&nbsp;</td>
		<td class="headrow" align="left">Version:</td>
		<td style="border: solid 1px #CCCCCC" width="170"><%=(rsIssue.Fields.Item("ncvVersion").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" align="left">Tested:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvTested").Value)%>&nbsp;</td>
		<td class="headrow" align="left">Approved:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvApproved").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" align="left" nowrap>Assigned To:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvAsigned_to_Orig").Value)%>&nbsp;</td>
		<td class="headrow" align="left" nowrap>Currently Assigned To:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvAsigned_to").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" align="left">Priority:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvPriority").Value)%>&nbsp;</td>
		<td class="headrow" align="left">Status:</td>
		<td style="border: solid 1px #CCCCCC"><%=(rsIssue.Fields.Item("chvStatus").Value)%>&nbsp;</td>
    </tr>
    <tr> 
		<td class="headrow" valign="top" align="left">Notes:</td>
		<td colspan="3"><textarea cols="75" rows="7" readonly><%=(rsIssue.Fields.Item("chvDescription").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top">Response:</td>
		<td nowrap valign="top"><textarea name="Response" cols="75" rows="7" tabindex="1" accesskey="F"></textarea></td>
    </tr>
    <tr> 
		<td nowrap>Current Status:</td>
		<td nowrap><select name="CurrentStatus" tabindex="2">
			<% 
			while (!rsStatus.EOF) { 
			%>
				<option value="<%=(rsStatus.Fields.Item("intStatus_id").Value)%>" <%=((rsStatus.Fields.Item("intStatus_id").Value==rsIssue.Fields.Item("intStatus_id").Value)?"SELECTED":"")%>><%=(rsStatus.Fields.Item("ncvStatus").Value)%></option>
			<% 
				rsStatus.MoveNext();
			}
			%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Version:</td>
		<td nowrap><input type="text" name="Version" value="<%=(rsIssue.Fields.Item("ncvVersion").Value)%>" size="8" tabindex="3"></td>
    </tr>
    <tr> 
		<td nowrap>Approved:</td>
		<td nowrap><input type="checkbox" name="Approved" CHECKED tabindex="4" class="chkstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Tested:</td>
		<td nowrap><input type="checkbox" name="Tested" tabindex="5" class="chkstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Assigned To:</td>
		<td nowrap><select name="AssignedTo" tabindex="6">
			<% 
			while (!rsStaff.EOF) { 
			%>
				<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==rsIssue.Fields.Item("intAssigned_to_orig").Value)?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvname").Value)%></option>
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
		<td><input type="button" value="Save" tabindex="7" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="9" onClick="history.back();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="FunctionID" value="<%=(rsIssue.Fields.Item("insFTNid").Value)%>">
<input type="hidden" name="Priority" value="<%=(rsIssue.Fields.Item("intPriority_id").Value)%>">
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>
<%
rsStaff.Close();
rsStatus.Close();
rsIssue.Close();
%>