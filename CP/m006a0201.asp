<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_NewAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_NewAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var rsCompanyAddress = Server.CreateObject("ADODB.Recordset");
	rsCompanyAddress.ActiveConnection = MM_cnnASP02_STRING;
	var StreetAddress = String(Request.Form("StreetAddress")).replace(/'/g, "''");			
	var City = String(Request.Form("City")).replace(/'/g, "''");			
	rsCompanyAddress.Source = "{call dbo.cp_company_address(0,0,'"+StreetAddress+"','"+City+"',"+Request.Form("Province")+",'"+Request.Form("PostalCode")+"',"+Request.Form("PrimaryPhoneType")+",'"+Request.Form("PrimaryPhoneAreaCode")+"','"+Request.Form("PrimaryPhoneNumber")+"','"+Request.Form("PrimaryPhoneExtension")+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Request.Form("SecondaryPhoneNumber")+"','"+Request.Form("SecondaryPhoneExtension")+"',0,'','','','"+Request.Form("Email")+"','',9,0,'E',0)}";
	rsCompanyAddress.CursorType = 0;
	rsCompanyAddress.CursorLocation = 2;
	rsCompanyAddress.LockType = 3;
	rsCompanyAddress.Open();
	Response.Redirect("InsertSuccessful.html");	
}

var rsProvince = Server.CreateObject("ADODB.Recordset");
rsProvince.ActiveConnection = MM_cnnASP02_STRING;
rsProvince.Source = "{call dbo.cp_Prov_State}";
rsProvince.CursorType = 0;
rsProvince.CursorLocation = 2;
rsProvince.LockType = 3;
rsProvince.Open();

var rsPhoneType = Server.CreateObject("ADODB.Recordset");
rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
rsPhoneType.Source = "{call dbo.cp_Phone_Type}";
rsPhoneType.CursorType = 0;
rsPhoneType.CursorLocation = 2;
rsPhoneType.LockType = 3;
rsPhoneType.Open();

var rsAreaCode = Server.CreateObject("ADODB.Recordset");
rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
rsAreaCode.Source = "{call dbo.cp_area_code(0,'',0,2,'Q',0)}";
rsAreaCode.CursorType = 0;
rsAreaCode.CursorLocation = 2;
rsAreaCode.LockType = 3;
rsAreaCode.Open();
%>
<html>
<head>
	<title>New Address</title>
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
			document.frm0201.reset();
			break;
	   	case 76 :
			//alert("L");
			window.location.href='m006q0201.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckPostalCode(document.frm0201.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm0201.PostalCode.focus();
			return ;
		}
		if (!CheckEmail(document.frm0201.EMail.value)){
			alert("Invalid Email.");
			document.frm0201.EMail.focus();
			return ;
		}
		var tempPC = document.frm0201.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm0201.PostalCode.value = tempPC;		
				
		document.frm0201.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0201.StreetAddress.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>New Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="1" accesskey="F"></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" maxlength="50" tabindex="2"></td>
	</tr>
	<tr> 
		<td nowrap>Province:</td>
		<td nowrap><select name="Province" tabindex="3">
			<% 
			while (!rsProvince.EOF) {
			%>
				<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>"><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%>
			<%
				rsProvince.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr>
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" tabindex="4" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr>
		<td nowrap>Primary Phone:</td>
		<td nowrap>
			<select name="PrimaryPhoneType" tabindex="5">
			<%
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>"><%=(rsPhoneType.Fields.Item("chvName").Value)%>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="6">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>"><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" size="9" tabindex="7" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">
			Ext <input type="text" name="PrimaryPhoneExtension" size="4" tabindex="8" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap>
			<select name="SecondaryPhoneType" tabindex="9">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value == rsCompanyAddress.Fields.Item("intPhone_Type_2").Value)?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="10">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>"><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" size="9" tabindex="11" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">
			Ext <input type="text" name="SecondaryPhoneExtension" size="4" tabindex="12" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" tabindex="13" accesskey="L"></td>
	</tr>  
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="14" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="15" onClick="window.location.href='m006e0101.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>