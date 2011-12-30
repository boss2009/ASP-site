<%@language="JAVASCRIPT"%> 
<!--#INCLUDE file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var OrganizationName = String(Request.Form("OrganizationName")).replace(/'/g, "'");		
	if (String(Request.Form("OrganizationType"))=="13") {
		var IsVendor = ((Request.Form("IsVendor")=="1")?"1":"0");	
		var IsManufacturer = ((Request.Form("IsManufacturer")=="1")?"1":"0");	
		var IsServiceProvider =	((Request.Form("IsServiceProvider")=="1")?"1":"0");		
	} else {
		var IsVendor = 0;
		var IsManufacturer = 0;
		var IsServiceProvider = 0;
	}
	var cmdInsertCompany = Server.CreateObject("ADODB.Command");
	cmdInsertCompany.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertCompany.CommandText = "dbo.cp_Company2";
	cmdInsertCompany.CommandType = 4;
	cmdInsertCompany.CommandTimeout = 0;
	cmdInsertCompany.Prepared = true;
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@chvCompany_Name", 200, 1,50,OrganizationName));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@insWork_Typ_id", 2, 1,1,Request.Form("OrganizationType")));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@bitIs_Vendor", 2, 1,1,IsVendor));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@bitIs_Mnfucter", 2, 1,1,IsManufacturer));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@bitIs_SrvPvdr", 2, 1,1,IsServiceProvider));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@insUser_id", 2, 1,2,Session("insStaff_id")));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@inspSrtBy", 2, 1,2,0));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@inspSrtOrd", 2, 1,2,0));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@chvFilter", 200, 1,50,""));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@chvTask", 129, 1,1,"A"));
	cmdInsertCompany.Parameters.Append(cmdInsertCompany.CreateParameter("@intRtnFlag", 3, 2));
	cmdInsertCompany.Execute();
			
	Response.Redirect("m006FS3.asp?intCompany_id="+cmdInsertCompany.Parameters.Item("@intRtnFlag").Value);	
}

var rsType = Server.CreateObject("ADODB.Recordset");
rsType.ActiveConnection = MM_cnnASP02_STRING;
rsType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
rsType.CursorType = 0;
rsType.CursorLocation = 2;
rsType.LockType = 3;
rsType.Open();
%>
<html>
<head>
	<title>New Organization</title>
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
		if (Trim(document.frm0101.OrganizationName.value)==""){
			alert("Enter Organization Name.");
			document.frm0101.OrganizationName.focus();
			return ;
		}
		document.frm0101.submit();
	}
	
	function ChangeType(){
		if (document.frm0101.OrganizationType.value=="13") {
			TypeCompany.style.visibility="visible";
		} else {
			TypeCompany.style.visibility="hidden";
		}
	}
	
	function Init(){
		ChangeType();
		document.frm0101.OrganizationName.focus();
	}
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_updateAction%>" method="POST" name="frm0101">
<h5>New Organization</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Organization Name:</td>
		<td nowrap><input type="text" name="OrganizationName" maxlength="50" size="50" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Organization Type:</td>
		<td nowrap><select name="OrganizationType" tabindex="2" onChange="ChangeType();">
			<%
			while (!rsType.EOF){
			%>
				<option value="<%=(rsType.Fields.Item("intWork_type_id").Value)%>"><%=(rsType.Fields.Item("chvWork_type_desc").Value)%> 
			<%
				rsType.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap colspan="2"><div id="TypeCompany" style="visibility: visible"> 
			<input type="checkbox" name="IsVendor" value="1" tabindex="3" class="chkstyle">Is Vendor 
			<input type="checkbox" name="IsManufacturer" value="1" tabindex="4" class="chkstyle">Is Manufacturer 
			<input type="checkbox" name="IsServiceProvider" value="1" tabindex="5" accesskey="L" class="chkstyle">Is Service Provider
		</div></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="7" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsType.Close()
%>