<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var LoanCategory = ((Request.Form("LoanCategory")=="1") ? "1":"0");
	var GrantCategory = ((Request.Form("GrantCategory")=="1") ? "1":"0");
	var OtherCategory = ((Request.Form("OtherCategory")=="1") ? "1":"0");			
	var ClientModule = ((Request.Form("ClientModule")=="1") ? "1":"0");		
	var LoanModule = ((Request.Form("LoanModule")=="1") ? "1":"0");
	var BuyoutModule = ((Request.Form("BuyoutModule")=="1") ? "1":"0");		
	var InstitutionModule = ((Request.Form("InstitutionModule")=="1") ? "1":"0");				
	var rsFundingSource = Server.CreateObject("ADODB.Recordset");
	rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsFundingSource.Source = "{call dbo.cp_Funding_Source_attributes("+ Request.Form("MM_recordId") + "," + LoanCategory + "," + GrantCategory + ","+ OtherCategory + ","+ ClientModule + ","+ LoanModule + "," + BuyoutModule + "," + InstitutionModule + ",0,'E',0)}";
	rsFundingSource.CursorType = 0;
	rsFundingSource.CursorLocation = 2;
	rsFundingSource.LockType = 3;
	rsFundingSource.Open();
	Response.Redirect("m018q03126.asp");
}

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source_attributes("+Request.QueryString("insFunding_source_id")+",0,0,0,0,0,0,0,1,'Q',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();
%>
<html>
<head>
	<title>Update Funding Source Attribute Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm03126.reset();
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
		if (Trim(document.frm03126.Description.value)==""){
			alert("Enter Description.");
			document.frm03126.Description.focus();
			return ;		
		}
		document.frm03126.submit();
	}
	</script>
</head>
<body onLoad="document.frm03126.Description.focus();">
<form name="frm03126" method="POST" action="<%=MM_editAction%>">
<h5>Update Funding Source Attribute Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%>" tabindex="1" readonly accesskey="F"></td>
    </tr>
	<tr>
		<td nowrap>Category:</td>
		<td></td>
	</tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="LoanCategory" <%=((rsFundingSource.Fields.Item("bitIs_Loan_Catagory").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" class="chkstyle">Loan</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="GrantCategory" <%=((rsFundingSource.Fields.Item("bitIs_Grant_Catagory").Value == 1)?"CHECKED":"")%> value="1" tabindex="3" class="chkstyle">Grant</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="OtherCategory" <%=((rsFundingSource.Fields.Item("bitIs_Other_Catagory").Value == 1)?"CHECKED":"")%> value="1" tabindex="4" class="chkstyle">Other</td>
    </tr>
	<tr>
		<td nowrap>Module:</td>
		<td></td>
	</tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="ClientModule" <%=((rsFundingSource.Fields.Item("bitIs_Adult_Client_Mod").Value == 1)?"CHECKED":"")%> value="1" tabindex="5" class="chkstyle">Client</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="LoanModule" <%=((rsFundingSource.Fields.Item("bitIs_Loan_Mod").Value == 1)?"CHECKED":"")%> value="1" tabindex="6" class="chkstyle">Loan</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="BuyoutModule" <%=((rsFundingSource.Fields.Item("bitIs_Buyout_Mod").Value == 1)?"CHECKED":"")%> value="1" tabindex="7" class="chkstyle">Buyout</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="InstitutionModule" <%=((rsFundingSource.Fields.Item("bitIs_School_Mod").Value == 1)?"CHECKED":"")%> value="1" tabindex="8" accesskey="L" class="chkstyle">Institution</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="9" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="10" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="11" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsFundingSource.Fields.Item("insFunding_source_id").Value %>">
</form>
</body>
</html>
<%
rsFundingSource.Close();
%>