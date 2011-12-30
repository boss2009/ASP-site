<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_Insert"))=="true") {
	var rsNewAddress__chrAddress_type = ((String(Request.Form("AddressType")) == "Permanent")?"P":"A");
	var Address = String(Request.Form("StreetAddress")).replace(/'/g, "''");
	var City = String(Request.Form("City")).replace(/'/g, "''");
	var PostalZip = String(Request.Form("PostalCode")).replace(/'/g, "''");
	var Email = String(Request.Form("EMail"));
	var Notes = String(Request.Form("Notes"));
	
	var rsNewAddress = Server.CreateObject("ADODB.Recordset");
	rsNewAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsNewAddress.Source="{call dbo.cp_Insert_Adult_Address("+Request.Form("intAdult_id")+",'"+rsNewAddress__chrAddress_type+"','"+Address+"','"+City+ "',"+Request.Form("Province")+",'"+PostalZip+"',"+Request.Form("PrimaryPhoneType")+",'"+Request.Form("PrimaryPhoneAreaCode") + "','"+ Request.Form("PrimaryPhoneNumber") + "','"+ Request.Form("PrimaryPhoneExtension") + "',"+ Request.Form("SecondaryPhoneType") + ",'"+ Request.Form("SecondaryPhoneAreaCode")+ "','"+Request.Form("SecondaryPhoneNumber")+"','"+Request.Form("SecondaryPhoneExtension") + "',0,'','','','"+ Email + "','"+ Notes.replace(/'/g, "''") + "',1,0)}";
	rsNewAddress.CursorType = 0;
	rsNewAddress.CursorLocation = 2;
	rsNewAddress.LockType = 3;
	rsNewAddress.Open();
	
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
	<title>New <%=((Request.QueryString("AddressType")=="Permanent")?"Permanent Address":"Address While Attending School")%></title>
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
				self.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function CopyFields(){
		document.frm0301.MM_Insert.value="false";
		document.frm0301.MM_CopyFields.value="true";
		document.frm0301.submit();
	}
		
	function Save(){
		if (!CheckTextArea(document.frm0301.Notes, 50)){
			alert("Text area cannot exceed 50 characters.");
			return ;
		}	
		if (!CheckPostalCode(document.frm0301.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm0301.PostalCode.focus();
			return ;
		}
		if (!CheckEmail(document.frm0301.EMail.value)){
			alert("Invalid Email.");
			document.frm0301.EMail.focus();
			return ;
		}
		var tempPC = document.frm0301.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm0301.PostalCode.value = tempPC;		
		document.frm0301.submit();
	}

	function Init(){
	<%
	var rsClientAddress__intpAdult_id = String(Request.QueryString("intAdult_id"));
	var rsClientAddress = Server.CreateObject("ADODB.Recordset");
	rsClientAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsClientAddress.Source = "{call dbo.cp_Adult_Address("+ rsClientAddress__intpAdult_id.replace(/'/g, "''") + ")}";
	rsClientAddress.CursorType = 0;
	rsClientAddress.CursorLocation = 2;
	rsClientAddress.LockType = 3;
	rsClientAddress.Open();
	
	if (rsClientAddress.EOF) {
	%>
		document.frm0301.Copy.disabled = true;	
	<%
	}
	if (String(Request.Form("MM_CopyFields"))=="true") {
		var rsAddress__intpAddr_id = "";
		if (!rsClientAddress.EOF) { 
			rsAddress__intpAddr_id = String(rsClientAddress.Fields.Item("intaddr_id").Value);
			var rsAddress = Server.CreateObject("ADODB.Recordset");
			rsAddress.ActiveConnection = MM_cnnASP02_STRING;
			rsAddress.Source = "{call dbo.cp_Idv_Adult_Address("+ rsAddress__intpAddr_id.replace(/'/g, "''") + ")}";
			rsAddress.CursorType = 0;
			rsAddress.CursorLocation = 2;
			rsAddress.LockType = 3;
			rsAddress.Open();
	%>
		document.frm0301.StreetAddress.value="<%=(rsAddress.Fields.Item("chvAddress").Value)%>";
		document.frm0301.City.value="<%=(rsAddress.Fields.Item("chvCity").Value)%>";
		document.frm0301.Province.value="<%=rsAddress.Fields.Item("insProv_State_id").Value%>";
		document.frm0301.PostalCode.value="<%=(rsAddress.Fields.Item("chvPostal_zip").Value)%>";
		document.frm0301.PrimaryPhoneType.value="<%=rsAddress.Fields.Item("intPhone_Type_1").Value%>";
		document.frm0301.PrimaryPhoneAreaCode.value="<%=rsAddress.Fields.Item("chvPhone1_Arcd").Value%>";
		document.frm0301.PrimaryPhoneNumber.value="<%=(rsAddress.Fields.Item("chvPhone1_Num").Value)%>";
		document.frm0301.PrimaryPhoneExtension.value="<%=(rsAddress.Fields.Item("chvPhone1_Ext").Value)%>";
		document.frm0301.SecondaryPhoneType.value="<%=rsAddress.Fields.Item("intPhone_Type_2").Value%>";
		document.frm0301.SecondaryPhoneAreaCode.value="<%=rsAddress.Fields.Item("chvPhone2_Arcd").Value%>";
		document.frm0301.SecondaryPhoneNumber.value="<%=(rsAddress.Fields.Item("chvPhone2_Num").Value)%>";
		document.frm0301.SecondaryPhoneExtension.value="<%=(rsAddress.Fields.Item("chvPhone2_Ext").Value)%>";
		document.frm0301.EMail.value="<%=(rsAddress.Fields.Item("chvEmail").Value)%>";
		document.frm0301.Notes.value="<%=(rsAddress.Fields.Item("chvNote").Value)%>";	
	<%
		}
	}
	%>
		document.frm0301.StreetAddress.focus();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0301">
<h5>New <%=((String(Request.QueryString("AddressType"))=="Permanent")?"Permanent Address":"Address While Attending School")%></h5>
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
			<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>" <%=((rsProvince.Fields.Item("intprvst_id").Value=="1")?" SELECTED":"")%>><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%></option>
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
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value=="1")?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
	        </select>
			<select name="PrimaryPhoneAreaCode" tabindex="6">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsAreaCode.Fields.Item("chvAC_num").Value=="604")?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>
	        </select>
			<input type="text" name="PrimaryPhoneNumber" size="9" tabindex="7" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext
			<input type="text" name="PrimaryPhoneExtension" size="4" tabindex="8" onKeypress="AllowNumericOnly();" maxlength="4">
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap> 
			<select name="SecondaryPhoneType" tabindex="9">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>"><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
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
    	    <input type="text" name="SecondaryPhoneNumber" size="9" tabindex="11" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext
			<input type="text" name="SecondaryPhoneExtension" size="4" tabindex="12" onKeypress="AllowNumericOnly();" maxlength="4" >
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" tabindex="13" size="20"></td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="14"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" name="Copy" value="Copy From <%=((String(Request.QueryString("AddressType"))=="Permanent")?"Alternative":"Permanent")%>" tabindex="16" onClick="CopyFields();" class="btnstyle"></td>		
		<td><input type="button" value="Close" tabindex="17" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="intAdult_id" value="<%=Request.QueryString("intAdult_id")%>">
<input type="hidden" name="MM_Insert" value="true">
<input type="hidden" name="MM_CopyFields" value="false">
<input type="hidden" name="AddressType" value="<%=Request.QueryString("AddressType")%>">
</form>
</body>
</html>
<%
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>