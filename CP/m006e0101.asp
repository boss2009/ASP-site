<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var OrganizationName = String(Request.Form("OrganizationName")).replace(/'/g, "''");		
	if (String(Request.Form("OrganizationType"))=="13") {
		var IsVendor = ((Request.Form("IsVendor")=="1")?"1":"0");	
		var IsManufacturer = ((Request.Form("IsManufacturer")=="1")?"1":"0");	
		var IsServiceProvider =	((Request.Form("IsServiceProvider")=="1")?"1":"0");		
	} else {
		var IsVendor = 0;
		var IsManufacturer = 0;
		var IsServiceProvider = 0;
	}
	var rsOrganization = Server.CreateObject("ADODB.Recordset");
	rsOrganization.ActiveConnection = MM_cnnASP02_STRING;
	rsOrganization.Source = "{call dbo.cp_Company2(" + Request.Form("CompanyID") + ",'" + OrganizationName + "',"+Request.Form("OrganizationType")+","+IsVendor+","+IsManufacturer+","+IsServiceProvider+","+Session("insStaff_id")+",0,0,'',0,'E',0)}";
	rsOrganization.CursorType = 0;
	rsOrganization.CursorLocation = 2;
	rsOrganization.LockType = 3;
	rsOrganization.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m006e0101.asp&intCompany_id="+Request.Form("CompanyID"));
}

var rsOrganization = Server.CreateObject("ADODB.Recordset");
rsOrganization.ActiveConnection = MM_cnnASP02_STRING;
rsOrganization.Source = "{call dbo.cp_Company2("+Request.QueryString("intCompany_id")+",'',0,0,0,0,0,1,0,'',1,'Q',0)}"
rsOrganization.CursorType = 0;
rsOrganization.CursorLocation = 2;
rsOrganization.LockType = 3;
rsOrganization.Open();	

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
	<title>General Information</title>
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
			TypeOrganization.style.visibility="visible";
		} else {
			TypeOrganization.style.visibility="hidden";
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
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Organization Name:</td>
		<td nowrap><input type="text" name="OrganizationName" maxlength="50" value="<%=rsOrganization.Fields.Item("chvCompany_Name").Value%>" size="50" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Organization Type:</td>
		<td nowrap><select name="OrganizationType" tabindex="2" onChange="ChangeType();">
			<%
			while (!rsType.EOF){
			%>
				<option value="<%=(rsType.Fields.Item("intWork_type_id").Value)%>" <%=((rsOrganization.Fields.Item("insWork_Typ_id").Value==rsType.Fields.Item("intWork_type_id").Value)?"SELECTED":"")%>><%=(rsType.Fields.Item("chvWork_type_desc").Value)%> 
			<%
				rsType.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap colspan="2"><div id="TypeOrganization" style="visibility: visible">
			<input type="checkbox" name="IsVendor" <%=((rsOrganization.Fields.Item("bitIs_Vendor").Value=="1")?"CHECKED":"")%> value="1" tabindex="3" class="chkstyle">Is Vendor
			<input type="checkbox" name="IsManufacturer" <%=((rsOrganization.Fields.Item("bitIs_Mnfucter").Value=="1")?"CHECKED":"")%> value="1" tabindex="4" class="chkstyle">Is Manufacturer
			<input type="checkbox" name="IsServiceProvider" <%=((rsOrganization.Fields.Item("bitIs_SrvPvdr").Value=="1")?"CHECKED":"")%> value="1" tabindex="5" accesskey="L" class="chkstyle">Is Service Provider
		</div></td>		
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="7" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="8" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="CompanyID" value="<%=Request.QueryString("intCompany_id")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsOrganization.Close()
rsType.Close()
%>