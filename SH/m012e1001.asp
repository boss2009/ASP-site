<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "update") {
	var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
	rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionAddress.Source = "{call dbo.cp_school_address(0,"+Request.Form("MM_recordId")+",'"+String(Request.Form("StreetAddress")).replace(/'/g, "''")+"','"+String(Request.Form("City")).replace(/'/g, "''")+"',"+Request.Form("ProvinceState")+",'"+Trim(Request.Form("PostalCode"))+"',"+Request.Form("PrimaryPhoneType")+",'"+Trim(Request.Form("PrimaryPhoneAreaCode"))+"','"+Trim(Request.Form("PrimaryPhoneNumber"))+"','"+Trim(Request.Form("PrimaryPhoneExtension"))+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Trim(Request.Form("SecondaryPhoneNumber"))+"','"+Trim(Request.Form("SecondaryPhoneExtension"))+"',0,'','','','"+Request.Form("Email")+"','',0,'E',0)}";
	rsInstitutionAddress.CursorType = 0;
	rsInstitutionAddress.CursorLocation = 2;
	rsInstitutionAddress.LockType = 3;
	rsInstitutionAddress.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012e1001.asp&insSchool_id="+Request.QueryString("insSchool_id"));	
}

if (String(Request("MM_action")) == "insert") {
	var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
	rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionAddress.Source = "{call dbo.cp_school_address("+Request.QueryString("insSchool_id")+",0,'"+String(Request.Form("StreetAddress")).replace(/'/g, "''")+"','"+String(Request.Form("City")).replace(/'/g, "''")+"',"+Request.Form("ProvinceState")+",'"+Request.Form("PostalCode")+"',"+Request.Form("PrimaryPhoneType")+",'"+Request.Form("PrimaryPhoneAreaCode")+"','"+Request.Form("PrimaryPhoneNumber")+"','"+Request.Form("PrimaryPhoneExtension")+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Request.Form("SecondaryPhoneNumber")+"','"+Request.Form("SecondaryPhoneExtension")+"',0,'','','','"+Request.Form("Email")+"','',0,'A',0)}";
	rsInstitutionAddress.CursorType = 0;
	rsInstitutionAddress.CursorLocation = 2;
	rsInstitutionAddress.LockType = 3;
	rsInstitutionAddress.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012e1001.asp&insSchool_id="+Request.QueryString("insSchool_id"));	
}

var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
rsInstitutionAddress.Source = "{call dbo.cp_school_address("+ Request.QueryString("insSchool_id") + ",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',1,'Q',0)}";
rsInstitutionAddress.CursorType = 0;
rsInstitutionAddress.CursorLocation = 2;
rsInstitutionAddress.LockType = 3;
rsInstitutionAddress.Open();

var intAddress_id = 0;
var IsNew = false;
if (rsInstitutionAddress.EOF) {
	IsNew = true;
} else {
	if ((rsInstitutionAddress.Fields.Item("intAddress_id").Value == 0) || (rsInstitutionAddress.Fields.Item("intAddress_id").Value == null)){
		IsNew = true;
	} else {
		intAddress_id = rsInstitutionAddress.Fields.Item("intAddress_id").Value;
		IsNew = false;
	}
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
	<title>Update Address</title>
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
				document.frm1001.reset();
			break;
		   	case 76 :
				//alert("L");
				window.location.href='m012e0101.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
		document.frm1001.StreetAddress.focus();
	}
	
	function Save(){
		if (!CheckPostalCode(document.frm1001.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm1001.PostalCode.focus();
			return ;
		}
		if (!CheckEmail(document.frm1001.EMail.value)){
			alert("Invalid Email.");
			document.frm1001.EMail.focus();
			return ;
		}
		var tempPC = document.frm1001.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm1001.PostalCode.value = tempPC;					
		document.frm1001.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm1001">
<h5>Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="1" accesskey="F"><%=((!IsNew)?rsInstitutionAddress.Fields.Item("chvAddress").Value:"")%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=((!IsNew)?rsInstitutionAddress.Fields.Item("chvCity").Value:"")%>" maxlength="50" tabindex="2" ></td>
	</tr>
	<tr> 
		<td nowrap>Province/State:</td>
		<td nowrap><select name="ProvinceState" tabindex="3">
		<% 
		while (!rsProvince.EOF) {
		%>
			<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>" <%if (!IsNew) {Response.Write(((rsProvince.Fields.Item("intprvst_id").Value==rsInstitutionAddress.Fields.Item("intprvst_id").Value)?"SELECTED":"")) } else {Response.Write(((rsProvince.Fields.Item("intprvst_id").Value==1)?"SELECTED":""))}%>><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%></option>
		<%
			rsProvince.MoveNext();
		}
		%>
		</select></td>
    </tr>	
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><input type="text" name="Country" value="<%=((!IsNew)?rsInstitutionAddress.Fields.Item("chvcntry_name").Value:"")%>" tabindex="4" readonly></td>
	</tr>
    <tr> 
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" value="<%=((!IsNew)?FormatPostalCode(rsInstitutionAddress.Fields.Item("chvPostal_zip").Value):"")%>" tabindex="5" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap> 
			<select name="PrimaryPhoneType" tabindex="6">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsPhoneType.Fields.Item("intPhone_type_id").Value==rsInstitutionAddress.Fields.Item("intPhone_Type_1").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="7">
				<option value="" <%if (!IsNew) Response.Write(((rsInstitutionAddress.Fields.Item("chvPhone1_Arcd").Value=="")?"SELECTED":""))%>>			
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsAreaCode.Fields.Item("chvAC_num").Value==rsInstitutionAddress.Fields.Item("chvPhone1_Arcd").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsInstitutionAddress.Fields.Item("chvPhone1_Num").Value):"")%>" size="9" tabindex="8" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)" >
			Ext <input type="text" name="PrimaryPhoneExtension" value="<%=((!IsNew)?rsInstitutionAddress.Fields.Item("chvPhone1_Ext").Value:"")%>" size="4" tabindex="9" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap>
			<select name="SecondaryPhoneType" tabindex="10">
			<% 
			rsPhoneType.MoveFirst();
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsPhoneType.Fields.Item("intPhone_type_id").Value==rsInstitutionAddress.Fields.Item("intPhone_Type_2").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="11">
				<option value="" <%if (!IsNew) Response.Write(((rsInstitutionAddress.Fields.Item("chvPhone2_Arcd").Value=="")?"SELECTED":""))%>>
			<%
			rsAreaCode.MoveFirst();
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsAreaCode.Fields.Item("chvAC_num").Value==rsInstitutionAddress.Fields.Item("chvPhone2_Arcd").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsInstitutionAddress.Fields.Item("chvPhone2_Num").Value):"")%>" size="9" tabindex="12" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)" >
			Ext <input type="text" name="SecondaryPhoneExtension" value="<%=((!IsNew)?rsInstitutionAddress.Fields.Item("chvPhone2_Ext").Value:"")%>" size="4" tabindex="13" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" value="<%=((!IsNew)?Trim(rsInstitutionAddress.Fields.Item("chvEmail").Value):"")%>" tabindex="14" accesskey="L"></td>
	</tr>  
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="window.location.href='m012e0101.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=((!IsNew)?intAddress_id:"")%>">
</form>
</body>
</html>
<%
rsInstitutionAddress.Close();
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>