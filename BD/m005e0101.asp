<%@language="JAVASCRIPT"%> 
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_actionAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_actionAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_action")) == "Update"){
	var BundleName = String(Request.Form("BundleName")).replace(/'/g, "''");		
	ForCSG = ((Request.Form("CSG")=="1")?"1":"0");	
	ForLoan = ((Request.Form("Loan")=="1")?"1":"0");	
	var rsBundle = Server.CreateObject("ADODB.Recordset");
	rsBundle.ActiveConnection = MM_cnnASP02_STRING;
	rsBundle.Source = "{call dbo.cp_Bundle("+Request.QueryString("insBundle_id")+",'"+BundleName+"',"+Request.Form("BundleCost")+",2,"+ForCSG+","+ForLoan+",'"+Request.Form("Type")+"',"+Request.Form("Status")+","+Session("insStaff_id")+",0,0,'',0,'E',0)}"
	rsBundle.CursorType = 0;
	rsBundle.CursorLocation = 2;
	rsBundle.LockType = 3;
	rsBundle.Open();	
	Response.Redirect("UpdateSuccessful.asp");
}

if (String(Request.Form("MM_action")) == "Clone") {
	var BundleName = String(Request.Form("BundleName")).replace(/'/g, "''");		
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

	var rsComponent = Server.CreateObject("ADODB.Recordset");
	rsComponent.ActiveConnection = MM_cnnASP02_STRING;
	rsComponent.Source = "{call dbo.cp_bundle_eqp_class("+Request.QueryString("insBundle_id")+",0,0,'Q',0)}";
	rsComponent.CursorType = 0;
	rsComponent.CursorLocation = 2;
	rsComponent.LockType = 3;
	rsComponent.Open();

	var rsAddComponent = Server.CreateObject("ADODB.Recordset");
	rsAddComponent.ActiveConnection = MM_cnnASP02_STRING;
	rsAddComponent.CursorType = 0;
	rsAddComponent.CursorLocation = 2;
	rsAddComponent.LockType = 3;
	while (!rsComponent.EOF) {
		rsAddComponent.Source = "{call dbo.cp_bundle_eqp_class("+BundleID+","+rsComponent.Fields.Item("insEquip_Class_id").Value+",1,'A',0)}";
		rsAddComponent.Open();
		rsComponent.MoveNext();	
	}
	Response.Redirect("CloneSuccessful.asp?insBundle_id="+BundleID);
}

var rsBundle = Server.CreateObject("ADODB.Recordset");
rsBundle.ActiveConnection = MM_cnnASP02_STRING;
rsBundle.Source = "{call dbo.cp_Bundle("+Request.QueryString("insBundle_id")+",'',0.0,0,1,1,'',0,"+Session("insStaff_id")+",0,0,'',1,'Q',0)}"
rsBundle.CursorType = 0;
rsBundle.CursorLocation = 2;
rsBundle.LockType = 3;
rsBundle.Open();	
/*
var rsComponent = Server.CreateObject("ADODB.Recordset");
rsComponent.ActiveConnection = MM_cnnASP02_STRING;
rsComponent.Source = "{call dbo.cp_bundle_eqp_class("+Request.QueryString("insBundle_id")+",0,0,'Q',0)}";
rsComponent.CursorType = 0;
rsComponent.CursorLocation = 2;
rsComponent.LockType = 3;
rsComponent.Open();
var Total = 0;
while (!rsComponent.EOF) { 
	Total += rsComponent.Fields.Item("FltList_Unit_Cost").Value;
	rsComponent.MoveNext();
}
*/
%>
<html>
<head>
	<title><%=(rsBundle.Fields.Item("chvName").Value)%> - Bundle ID: <%=Request.QueryString("insBundle_id")%></title>
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
		document.frm0101.MM_action.value="Update";
		document.frm0101.submit();
	}
	
	function Clone1(){
		if (Trim(document.frm0101.BundleName.value)==""){
			alert("Enter Bundle Name.");
			document.frm0101.BundleName.focus();
			return ;
		}
		document.frm0101.MM_action.value="Clone";
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
				case "0":
					temp = temp + "Other ";
				break;				
			}
			document.frm0101.BundleName.value = temp + ": "+Trim(document.frm0101.BundleName.value);
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
				case "0":
					temp = temp + "Other ";
				break;
			}
			document.frm0101.BundleName.value = temp + ": " + Trim(NameSplit[1]);
		}
	}
	
	function Init(){
		ChangeBundleName();
		document.frm0101.BundleName.focus();
	}
	
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}		
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_actionAction%>" method="POST" name="frm0101">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Bundle Name:</td>
		<td nowrap><input type="text" name="BundleName" maxlength="50" value="<%=(rsBundle.Fields.Item("chvName").Value)%>" size="50" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Purpose:</td>
		<td nowrap>
			<input type="checkbox" name="CSG" value="1" <%=((rsBundle.Fields.Item("bitFor_CSG").Value=="1")?"checked":"")%> tabindex="2" onClick="ChangeBundleName();" class="chkstyle">CSG&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="Loan" value="1" <%=((rsBundle.Fields.Item("bitFor_Loan").Value=="1")?"checked":"")%> tabindex="3" onClick="ChangeBundleName();" class="chkstyle">Loan
		</td>
	</tr>
	<tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="4" onChange="ChangeBundleName();">
			<option value="1" <%=((rsBundle.Fields.Item("chrBundle_Type").Value=="1")?"SELECTED":"")%>>Desktop 
			<option value="2" <%=((rsBundle.Fields.Item("chrBundle_Type").Value=="2")?"SELECTED":"")%>>Laptop 
			<option value="0" <%=((rsBundle.Fields.Item("chrBundle_Type").Value=="0")?"SELECTED":"")%>>Other 
        </select></td>
	</tr>
	<tr> 
		<td nowrap>Status:</td>
		<td nowrap><select name="Status" tabindex="5">
			<option value="1" <%=((rsBundle.Fields.Item("bitBundle_Status").Value=="1")?"SELECTED":"")%>>Active 
			<option value="0" <%=((rsBundle.Fields.Item("bitBundle_Status").Value=="0")?"SELECTED":"")%>>Inactive 
        </select></td>
	</tr>
	<tr>
		<td nowrap>Bundle Cost:</td>
		<td nowrap>$<input type="text" name="BundleCost" size="8" value="<%=(rsBundle.Fields.Item("FltList_Unit_Cost").Value)%>" tabindex="6" accesskey="L"></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="7" class="btnstyle"></td>
		<td><input type="button" value="Save As New" onClick="Clone1();" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="9" class="btnstyle"></td>
		<td><input type="button" value="Add To Desktop" onClick="openWindow('m005a01j.asp?insBundle_id=<%=Request.QueryString("insBundle_id")%>','wj0501');" tabindex="10" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_action">
</form>
</body>
</html>
<%
rsBundle.Close();
%>