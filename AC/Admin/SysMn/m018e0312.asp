<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
if(String(Request.Form("MM_update")) == "true"){
	var MM_editRedirectUrl  = "m018e0312B.asp?insrefer_agent_id=" + Request.Form("MM_ID") + "&chvname=" + Request.Form("Description");	   
	MM_editRedirectUrl += ((Request.Form("IsActive")=="1")?"&bitis_active=1":"&bitis_active=0");
	MM_editRedirectUrl += ((Request.Form("IsLoan")=="1")?"&bitis_loan=1":"&bitis_loan=0");
	MM_editRedirectUrl += ((Request.Form("IsBuyout")=="1")?"&bitis_BuyOut=1":"&bitis_BuyOut=0");
    MM_editRedirectUrl += "&chrFS_chbx="   + Request.Form("MM_buffer");
	if (String(Request.Form("MM_buffer")) == "" ){
		Response.Write("<P>");	 
		Response.Write("No update is allowed if no check box is clicked ...");
		Response.Write("<P>");	 
	} else {
		Response.Redirect(MM_editRedirectUrl);
	}
}

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source(0,0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();

var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
rsReferringAgent.Source = "{call dbo.cp_referring_agent("+ Request.QueryString("insrefer_agent_id") + ",1,0,0,1)}";
rsReferringAgent.CursorType = 1;
rsReferringAgent.CursorLocation = 3;
rsReferringAgent.LockType = 3;
rsReferringAgent.Open();

if (String(rsReferringAgent.Fields.Item("chrFS_chbx").Value) != "null" ) {
	var chrFSOption = rsReferringAgent.Fields.Item("chrFS_chbx").Value;
	var FSArray = new Array();
	for (var i=0;i < chrFSOption.length; i++){
		FSArray[i] = chrFSOption.substr(i,1);
	}
} else {  
	Response.Write("No funding source."); 
}
%>
<html>
<head>
	<title>Update Referral Type Lookup</title>
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
				document.frm0312.reset();
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
		if (Trim(document.frm0312.Description.value)==""){
			alert("Enter Description.");
			document.frm0312.Description.focus();
			return ;
		}
		
		var buffer = "";   
		for (var i=0; i < document.frm0312.elements.length ; i++){
			if (document.frm0312.elements[i].name == "FundingSourceCheckBox") buffer += ((document.frm0312.elements[i].checked==1)?"1":"0");
		}
		document.frm0312.MM_buffer.value = buffer;
		document.frm0312.action = "m018e0312.asp";
		document.frm0312.submit();
	}
</Script>
</head>
<body onLoad="document.frm0312.Description.focus();">
<form name="frm0312" method="post" action="">
<h5>Update Referral Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsReferringAgent.Fields.Item("chvname").Value)%>" maxlength="40" size="40" accesskey="F" ></td>
    </tr>
    <tr> 
		<td nowrap>Is Active:</td>	
		<td nowrap><input type="checkbox" name="IsActive" <%=((rsReferringAgent.Fields.Item("bitis_active").Value == 1)?"CHECKED":"")%> value="1" class="chkstyle"></td>
	</tr>
	<tr>
		<td nowrap>Is Loan:</td> 
		<td nowrap><input type="checkbox" name="IsLoan" <%=((rsReferringAgent.Fields.Item("bitis_loan").Value == 1)?"CHECKED":"")%> value="1" class="chkstyle"></td>
    </tr>
	<tr>
		<td nowrap>Is Buyout:</td>
		<td nowrap><input type="checkbox" name="IsBuyOut" <%=((rsReferringAgent.Fields.Item("bitis_BuyOut").Value == 1)?"CHECKED":"")%> value="1" accesskey="L" class="chkstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Funding Source:</td>
		<td></td>
    </tr>
<% 
	var i = 0;
	while (!rsFundingSource.EOF) { 
%>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="FundingSourceCheckBox" <%=((FSArray[i] == 1)?"CHECKED":"")%> value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>" class="chkstyle"><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></td>
	</tr>
<%	
		i++;
		rsFundingSource.MoveNext();
	} 
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_buffer">
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_ID" value="<%=(rsReferringAgent.Fields.Item("insrefer_agent_id").Value)%>">
</form>
</body>
</html>
<%
rsReferringAgent.Close();
rsFundingSource.Close();
%>