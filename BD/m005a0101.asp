<%@language="JAVASCRIPT"%> 
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var BundleName = String(Request.Form("BundleName")).replace(/'/g, "'");		
	ForCSG = ((Request.Form("CSG")=="1")?"1":"0");	
	ForLoan = ((Request.Form("Loan")=="1")?"1":"0");	
	var cmdInsertBundle = Server.CreateObject("ADODB.Command");
	cmdInsertBundle.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertBundle.CommandText = "dbo.cp_Bundle";
	cmdInsertBundle.CommandType = 4;
	cmdInsertBundle.CommandTimeout = 0;
	cmdInsertBundle.Prepared = true;
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@chvBundle_Name", 200, 1,50,BundleName));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@FltList_Unit_Cost", 5, 1,1,Request.Form("BundleCost")));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@intPrice_Qty_id", 3, 1,1,1));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@bitFor_CSG", 2, 1,1,ForCSG));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@bitFor_Loan", 2, 1,1,ForLoan));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@chrBundle_Type", 129, 1,1,Request.Form("Type")));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@bitBundle_Status", 2, 1,1,Request.Form("Status")));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@inspSrtBy", 2, 1,1,0));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@chvFilter", 200, 1,1,''));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertBundle.Parameters.Append(cmdInsertBundle.CreateParameter("@intRtnFlag", 3, 2));
	cmdInsertBundle.Execute();
	
	var BundleID = cmdInsertBundle.Parameters.Item("@intRtnFlag").Value;
	
	Response.Redirect("m005FS3.asp?insBundle_id="+BundleID);	
}
%>
<html>
<head>
	<title>New Equipment Bundle</title>
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
				document.frm0101.reset();
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
		if (Trim(document.frm0101.BundleName.value)==""){
			alert("Enter Bundle Name.");
			document.frm0101.BundleName.focus();
			return ;
		}

		if (isNaN(document.frm0101.BundleCost.value)){
			alert("Invaild Bundle Cost.");
			document.frm0101.BundleCost.focus();
			return ;
		}		
		document.frm0101.submit();
	}
	
	function ChangeBundleName(){
		var temp = "";
		var NameSplit;
		NameSplit = document.frm0101.BundleName.value.split(":");
		if (NameSplit.length==1) {
			if (document.frm0101.CSG.checked && !document.frm0101.Loan.checked) temp = "CSG ";
			if (document.frm0101.Loan.checked && !document.frm0101.CSG.checked) temp = "Loan ";
			if (document.frm0101.CSG.checked && document.frm0101.Loan.checked) temp = "CSG/Loan ";
			switch (document.frm0101.Type.value) {
				case "1":
					temp = temp + "Desktop ";
				break;
				case "2":
					temp = temp + "Laptop ";
				break;
			}
			document.frm0101.BundleName.value = temp + ": " +Trim(document.frm0101.BundleName.value);
		} else {
			if (document.frm0101.CSG.checked && !document.frm0101.Loan.checked) temp = "CSG ";
			if (document.frm0101.Loan.checked && !document.frm0101.CSG.checked) temp = "Loan ";
			if (document.frm0101.CSG.checked && document.frm0101.Loan.checked) temp = "CSG/Loan ";
			switch (document.frm0101.Type.value) {
				case "1":
					temp = temp + "Desktop ";
				break;
				case "2":
					temp = temp + "Laptop ";
				break;
			}
			document.frm0101.BundleName.value = temp + ": " + Trim(NameSplit[1]);
		}
	}
	
	function Init(){
		ChangeBundleName();
		document.frm0101.BundleName.focus();
	}
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_updateAction%>" method="POST" name="frm0101">
<h5>New Equipment Bundle</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Bundle Name:</td>
		<td nowrap><input type="text" name="BundleName" maxlength="50" size="50" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Purpose:</td>
		<td nowrap>
			<input type="checkbox" name="CSG" value="1" tabindex="2" onClick="ChangeBundleName();" class="chkstyle">CSG
			<input type="checkbox" name="Loan" value="1" tabindex="3" onClick="ChangeBundleName();" class="chkstyle">Loan
		</td>
    </tr>
    <tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="4" onChange="ChangeBundleName();">
			<option value="1">Desktop 
			<option value="2">Laptop 
			<option value="0">Other 
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Status:</td>
		<td nowrap><select name="Status" tabindex="5">
			<option value="1">Active 
			<option value="0">Inactive 
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Bundle Cost:</td>
		<td nowrap>$<input type="text" name="BundleCost" value="0.00" onKeypress="AllowNumericOnly();" tabindex="6" accesskey="L"></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><input type="button" value="Save" onClick="Save();" tabindex="7" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="8" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>